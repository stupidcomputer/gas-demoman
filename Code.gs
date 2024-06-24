function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Demographic Manager')
    .addItem('Initialize Spreadsheet', 'spread_init')
    .addItem('Regenerate Reporting Spreadsheets', 'regenerate')
    .addToUi()
}

function spread_init() {
  var ui = SpreadsheetApp.getUi();
  // This is a destructive operation, so prompt the user.
  var response = ui.alert(
    'Initialize spreadsheet',
    'Warning: if you click yes, the spreadsheet will delete data inside itself. Make a backup before running this operation. Continue?',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response == ui.Button.NO) {
    ui.alert("Canceled the initialization of the spreadsheet.")
    return;
  } else if (response == ui.Button.CANCEL) {
    ui.alert("Canceled the initialization of the spreadsheet.")
    return;
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetsToRemove = spreadsheet.getSheets()
  for (var sheet = 0; sheet < sheetsToRemove.length; sheet++) {
    try {
      spreadsheet.deleteSheet(sheetsToRemove[sheet]);
    } catch { // ran out of spreadsheets to nuke, so rename it to the first thing we need
      spreadsheet.getSheets()[0].setName("*Demographic Data")
    }
  }

  // take care of demographic data layout
  var sheet = spreadsheet.getSheets()[0]
  sheet.getRange('A1').activate();
  sheet.getCurrentCell().setValue('Site Name');
  sheet.getRange('B1').activate();
  sheet.getCurrentCell().setValue('Count');
  sheet.getRange('C1').activate();
  sheet.getCurrentCell().setValue('Age');
  sheet.getRange('D1').activate();
  sheet.getCurrentCell().setValue('Gender');
  sheet.getRange('E1').activate();
  sheet.getCurrentCell().setValue('Ethnicity');
  sheet.getRange('F1').activate();
  sheet.getCurrentCell().setValue('Race');
  sheet.getRange('F2').activate();

  var newSheet = spreadsheet.insertSheet();
  newSheet.setName("*Site Information")
  newSheet.getRange('A1').activate();
  newSheet.getCurrentCell().setValue('Site Name');
  newSheet.getRange('B1').activate();
  newSheet.getCurrentCell().setValue('Date');
  newSheet.getRange('C1').activate();
  newSheet.getCurrentCell().setValue('Time');
  newSheet.getRange('D1').activate();
  newSheet.getCurrentCell().setValue('Those Present');
  newSheet.getRange('E1').activate();
  newSheet.getCurrentCell().setValue('Data Leads');
  newSheet.getRange('F1').activate();
  newSheet.getCurrentCell().setValue('Who Collected?');
  newSheet.getRange('F2').activate();
}