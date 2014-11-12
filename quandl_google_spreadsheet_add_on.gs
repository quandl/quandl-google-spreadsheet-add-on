/**
 * Quandl Import -- data retrieval into Google Spreadsheets from Quandl.com.
 *
 * For a given Quandl code, this will load and insert the dataset starting at the first unused row.
 *
 * Quandl codes can be passed in one of three ways:
 *
 * a) Enter the entire code into a single cell, highlight the cell, and choose
 *    Quandl/Read Quandl Data from the menu
 *
 * b) Enter the source code in one cell, then enter the table code in the immediately
 *    adjacent cell, highlight both cells, and then choose Quandl/Read Quandl Data from the menu
 *
 * c) Simply choose Quandl/Read Quandl Data from the menu and you will be asked for the code
 *
 * If the system cannot infer the code from highlighted cells, or cannot interpret a manually
 * entered code, the user will be asked to input the code (again) or cancel the operation
 *
 * Native Google Spreadsheet errors when a dataset can't be loaded are probably sufficient
 * in terms of error notifications for that condition.
 *
 */


/**
 * The main wrapper function. Called via menu selection.
 * Tries to determine the Quandl code, and auth token, and
 * then loads the data and appends it onto the current sheet.
 */
function readQuandlData() {

	// used to build the URL
	var protocol = "https"
	var domain = "www.quandl.com";
	var api_version = "1";
	var root_controller = "datasets";

	// auth_token is stored in the cache
	var cache = CacheService.getPrivateCache();
	var auth_token = cache.get('auth_token');
	var cache_expiration_in_seconds = 21600; //21600 seconds = 6 hours = maximum

	// only works with json right now
	var format = "json";

  var parameters = "?sort_order=desc";

  if (auth_token) {
		// replace the auth_token in the cache
		// this is done to continually push back the expiry
		cache.put('auth_token', auth_token, cache_expiration_in_seconds);
		parameters = parameters + "&auth_token=" + encodeURIComponent(auth_token);
	}

	// obtain a handle to the spreadsheet objects
	var spreadsheet_object = SpreadsheetApp;
	var active_sheet_object = spreadsheet_object.getActiveSheet();

	// get the active cell(s), and get the contents of those cell(s), i.e.: contents of a range
	var active_cell_contents = active_sheet_object.getActiveRange().getValues();

	// check to see if the contents contain a valid quandl code
	var quandl_code = getQuandlCodesFromRangeValues(active_cell_contents);

	if (!quandl_code) {
		return exitGracefully();
	}

	// construct the url 
	var full_url = protocol + "://" + domain + "/api/v" + api_version + "/" + root_controller + "/" + quandl_code + "." + format;

	// effect the call to the url
	var result_text = UrlFetchApp.fetch(full_url + parameters);

	if (format == "json") {

		// pull out the JSON content 
		var result_json = result_text.getContentText();

		// extract the required data into variables
		var column_names = JSON.parse(result_json)["column_names"];
		var data = JSON.parse(result_json)["data"];

		// append the column headers and the data
		active_sheet_object.appendRow(column_names);
		data.forEach(function(row) {
			active_sheet_object.appendRow(row);
		});

	} else {

		var ui = SpreadsheetApp.getUi();
		ui.alert("Non-JSON imports are not supported at this time.");

	}; // if format==json

	// return control to the user, data should be in the sheet

};

/**
 * A central function to tell the user that there has been an error and we can't continue
 * This will return control to the user without data being requested/presented.
 * This should be called only from the main wrapper function because that will
 * allow the script to end via a return of false.
 *
 */
function exitGracefully() {
	var ui = SpreadsheetApp.getUi();
	ui.alert("Quandl Import is not able to continue. Please try again.");
	return false;
}; //exitGracefully


/**
 * Analyses the passed value (a range object) and tries to determine if a source_code and table_code
 * are present.
 *
 * If it cannot, it asks the user for the information.
 *
 * Returns a valid code or false if the user cancels the operation.
 *
 */
function getQuandlCodesFromRangeValues(range_values) {

	// a variable for manipulation during the function.
	// ideally will contain a correct, valid code at the end of 
	// the function and will be returned.
	var quandl_code = null;

	// we can handle one selected row (range_values.length==1)
	if (range_values.length == 1) {

		// we determine how many cells are selected in this row
		if (range_values[0] instanceof Array) {
			number_of_selected_cells = range_values[0].length;
		} else {
			number_of_selected_cells = 1;
		}

		// we can handle one or two selected cells.
		if (number_of_selected_cells == 1) {
			// one cell is selected, assume the entire code is here
			quandl_code = range_values[0][0];
		} else if (number_of_selected_cells == 2) {
			// assume source_code in first cell and table_code in second cell
			// but we will put them back together here 
			quandl_code = range_values[0][0] + "/" + range_values[0][1];
		}; // number_of_selected_cells
	}; // active_cell_contents.length

	while (!isValidQuandlCode(quandl_code)) {
		quandl_code = requestCodeFromUser();
		if (!quandl_code) {
			break;
		}
	}; // while

	return quandl_code;

}; //getQuandlCodesFromRangeValues()


/**
 * Presents a prompt that asks the user to input
 * a Quandl code
 */
function requestCodeFromUser() {

	var ui = SpreadsheetApp.getUi();
	var response = ui.prompt('Quandl Code Required', 'Please enter the Quandl code in the correct format (e.g. TAMMER1/SHIBOR) for the dataset you wish to load and then 
click OK.', ui.ButtonSet.OK_CANCEL);
	var quandl_code = response.getResponseText();

	if (response.getSelectedButton() == ui.Button.OK && quandl_code) {
		return quandl_code;
	} else {
		// they cancelled or didn't answer
		return false;
	}; // response.getSelectedButton

}; //requestCodeFromUser


/**
 * Takes a string, and returns true or false.
 * Performs a simple cleanup and tests on the string to determine if it
 * can be used as a quandl code.
 */
function isValidQuandlCode(quandl_code) {

	// if we know its not valid, return immediately
	if (!quandl_code) {
		return false;
	}

	// A valid code is defined as a string that can be split on the slash (/)
	// and which returns at least two string elements of non-zero length
	// Some cleaning of the code is performed initially.

	// replace double-slashes with a single slash, in case they were repeated
	quandl_code = quandl_code.replace(/\/+/g, '/');

	// remove spaces in case they are introduced or leading/trailing from a copy/paste
	quandl_code = quandl_code.replace(/\s/g, '');

	// split the codes into the two segments on the slash
	var source_code = quandl_code.split("/")[0];
	var table_code = quandl_code.split("/")[1];

	// do we actually have data in each segment
	if ((source_code == null || table_code == null) || (source_code.length == 0 || table_code.length == 0)) {
		// not valid data, return false
		return false;
	}

	return true;

}; //isValidQuandlCode 


/**
 * Presents a prompt that asks the user to input
 * their authToken
 *
 */
function requestAuthTokenFromUser() {

	var ui = SpreadsheetApp.getUi();
	var response = ui.prompt('Quandl auth token required', 'Please enter your Quandl auth token. The system will store your auth token for future use.', 
ui.ButtonSet.OK_CANCEL);
	var auth_token = response.getResponseText();
	if (response.getSelectedButton() == ui.Button.OK && auth_token) {
		return auth_token;
	} else {
		// they cancelled or didn't answer
		return false;
	}; // response.getSelectedButton

}; //requestAuthTokenFromUser


/**
 * Calls a function to request the auth token
 * and then stores the token in the cache
 */
function enterAuthToken() {

	var auth_token = null;
	var cache = CacheService.getPrivateCache();
	var cache_expiration_in_seconds = 21600; //21600 seconds = 6 hours = maximum

	auth_token = requestAuthTokenFromUser();
	if (!auth_token) {
		return exitGracefully();
	};

	// put the auth_token into the cache
	cache.put('auth_token', auth_token, cache_expiration_in_seconds);

	var ui = SpreadsheetApp.getUi();
	ui.alert("Your auth token has been stored.");

}


/**
 * Clears the auth token from the private cache
 *
 */
function clearAuthToken() {
	var cache = CacheService.getPrivateCache();
	cache.remove('auth_token');
	var ui = SpreadsheetApp.getUi();
	if (!cache.get('auth_token')) {
		ui.alert("Your auth token has been cleared.");
	} else {
		ui.alert("There was an error clearing your auth token. Please try again later.");
	};
}; //clearAuthToken


/**
 * A Google Apps automatic function to populate the menu
 * with the Read Quandl Data option and Clear Auth Token option
 *
 */
function onOpen() {
	var active_sheet_object = SpreadsheetApp.getActiveSpreadsheet();
	var entries = [{
		name: "Read From Quandl Dataset",
		functionName: "readQuandlData"
	}, {
		name: "Enter Auth Token",
		functionName: "enterAuthToken"
	}, {
		name: "Clear Auth Token",
		functionName: "clearAuthToken"
	}];
	active_sheet_object.addMenu("Quandl", entries);
}; //onOpen
