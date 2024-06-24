function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:F1').activate();
  spreadsheet.getRange('A1:F1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};

function UntitledMacro2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setValue('Site Name');
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('Count');
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setValue('Age');
  spreadsheet.getRange('D1').activate();
  spreadsheet.getCurrentCell().setValue('Gender');
  spreadsheet.getRange('E1').activate();
  spreadsheet.getCurrentCell().setValue('Ethnicity');
  spreadsheet.getRange('F1').activate();
  spreadsheet.getCurrentCell().setValue('Race');
  spreadsheet.getRange('F2').activate();
};