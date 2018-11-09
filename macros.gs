function Createafilter() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:F').activate();
  spreadsheet.getRange('A1:F').createFilter();
};

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(0, 0, 54, 8).activate();
  spreadsheet.getActiveRange().createFilter();
};

function Turnofffilter() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().getFilter().remove();
};