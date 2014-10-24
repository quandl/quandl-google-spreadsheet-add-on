Quandl Google Spreadsheet Add On
================================

This is a simple tool to import Quandl datasets directly into Google Spreadsheets. It is written in Google Apps Script.

### For more information

- [Google Apps Script](https://developers.google.com/apps-script/)
- [Quandl](http://www.quandl.com/)

# Instructions

1. Download the file _quandl_google_spreadsheet_add_on.gs_ to your local machine and load it into an editor

2. Select and copy the text to your clipboard.

3. Create a new Google Spreadsheet

4. From the spreadsheet, choose the menu _Tools -> Script Editor_

5. In the next dialog box: _Create Script for -> Spreadsheet_

6. Paste the code into window

7. From the menu, choose _File -> Save_

8. Name the file, for example _quandl_importer_

9. Close the Script Editor tab and refresh your spreadsheet using F5 or CMD-R

10. You will see a new menu after "Help", called "Quandl"

# Example Use

1. In cell A1, enter a valid Quandl dataset code, for example: _ROSSBARCLAY/8OL_

2. Select cell A1

3. From the new Quandl menu, choose "Read from Quandl Dataset"

4. After you are asked for authorization, Click continue and Accept

5. Your data (in this example, Brazilian deforestation statistics) should appear starting on row B.

# Notes

* If you enter your Auth Token, it will remain valid for six hours after your last request.

* If you do not enter an Auth Token, the system will default to using no token. You may experience issues requesting excessive amounts of data without a valid token. 

