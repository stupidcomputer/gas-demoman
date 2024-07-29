var _demographic_data = "*Demographic Data";
var _site_information = "*Site Information";
var _intern_data      = "*Intern Data";

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
      spreadsheet.getSheets()[0].setName(_demographic_data)
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

  var newSheet = spreadsheet.insertSheet();
  newSheet.setName(_site_information)
  newSheet.getRange('A1').activate();
  newSheet.getCurrentCell().setValue('Site Name');
  newSheet.getRange('B1').activate();
  newSheet.getCurrentCell().setValue('Date');
  newSheet.getRange('C1').activate();
  newSheet.getCurrentCell().setValue('Those Present');
  newSheet.getRange('D1').activate();
  newSheet.getCurrentCell().setValue('Data Leads');
  newSheet.getRange('E1').activate();
  newSheet.getCurrentCell().setValue('On-Site Lead(s)');
  newSheet.getRange('F1').activate();
  newSheet.getCurrentCell().setValue('Who Collected?');

  newSheet = spreadsheet.insertSheet();
  newSheet.setName(_intern_data)
  newSheet.getRange('A1').activate();
  newSheet.getCurrentCell().setValue('First Name');
  newSheet.getRange('B1').activate();
  newSheet.getCurrentCell().setValue('Last Name');
  newSheet.getRange('C1').activate();
  newSheet.getCurrentCell().setValue('Gender');
  newSheet.getRange('D1').activate();
  newSheet.getCurrentCell().setValue('Ethnicity');
  newSheet.getRange('E1').activate();
  newSheet.getCurrentCell().setValue('Race');
}

function get_titled_spreadsheet_contents(sheet_name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var rows = spreadsheet.getDataRange().getValues();
  rows.shift(); /* remove the column headers from the data */

  return rows;
}

/*
 * @param {string} site_name
 */
function get_demographic_records_for_site(site_name) {
  var rows = get_titled_spreadsheet_contents(_demographic_data);
  var output = [];

  for(row of rows) {
    if(site_name != null && row[0] === site_name) {
      output = output.concat(Array(row[1]).fill({
        "site_name": row[0],
        "age": row[2],
        "gender": row[3],
        "ethnicity": row[4],
        "race": row[5],
      }));
    } else if(site_name === null) {
      output = output.concat(Array(row[1]).fill({
        "site_name": row[0],
        "age": row[2],
        "gender": row[3],
        "ethnicity": row[4],
        "race": row[5],
      }));
    }
  }

  return output;
}

function get_sites() {
  var rows = get_titled_spreadsheet_contents(_site_information);
  var output = [];

  for(row of rows) {
    output.push({
      "site_name": row[0],
      "date": row[1],
      "present": row[2],
      "data_leads": row[3],
      "on_site_leads": row[4],
      "compiled_by": row[5],
      "attached_records": [],
    })
  }

  return output;
}

function get_intern_data() {
  var rows = get_titled_spreadsheet_contents(_intern_data);
  var output = {};

  for(row of rows) {
    var no_last_name = row[1] === "";
    /* if there's no last name, then just use the first name
     * as the key */
    var key = row[0].concat(no_last_name ? "" : " ".concat(row[1]))
    output[key] = {
      "first_name": row[0],
      "last_name": row[1],
      "calculated_name": key,
      "age": "Intern",
      "gender": row[2],
      "ethnicity": row[3],
      "race": row[4],
    };
  }

  return output;
}

function collate_site_data() {
  var sites = get_sites();
  var intern_data = get_intern_data();

  for(site of sites) {
    var name = site["site_name"];
    var data = get_demographic_records_for_site(name);

    site["attached_records"].push(...data);
    var those_present = site["present"].split(',').map((x) => x.trim());

    for(person of those_present) {
      site["attached_records"].push(
        intern_data[person]
      )
    }
  }

  return sites;
}

function get_sums(data, field_name, human_readable) {
  var collated = {}
  
  for(datum of data) {
    var result = datum[field_name];

    if(datum[field_name] in collated) {
      collated[datum[field_name]] += 1
    } else {
      collated[datum[field_name]] = 1
    }
  }

  var output = [[
    human_readable, "Count"
  ]]

  for(const [key, value] of Object.entries(collated)) {
    output.push([String(key), String(value)]);
  }

  /* pad out the rest */
  var output_len = output.length;
  output = output.concat(Array(7 - output_len).fill(["", ""]))
  return output;
}

function filtration(data, interns, adults, children) {
  var output = [];

  for(datum of data) {
    if(datum === undefined) {
      SpreadsheetApp.getUi().alert("Couldn't find an intern -- skipping. See the manual for more information.")
      continue;
    }

    if(interns && datum.age == "Intern") {
      output.push(datum);
    } else if(adults && datum.age == "Adult") {
      output.push(datum);
    } else if(children && datum.age == "Child") {
      output.push(datum);
    }
  }

  return output;
}

function regenerate() {
  /* remove outdated sheets */
  var spreadsheet = SpreadsheetApp.getActive()
  var sheetList = spreadsheet.getSheets().map((x) => x.getName());

  for(sheet of sheetList) {
    if(sheet[0] === ">") {
      spreadsheet.deleteSheet(spreadsheet.getSheetByName(sheet));
    }
  }

  var sites = collate_site_data();
  for(site of sites) {
    var newSheet = spreadsheet.insertSheet(spreadsheet.getNumSheets());
    newSheet.setName(">".concat(site["site_name"]))

    /* create the header of the spreadsheet */
    newSheet.getRange('A1:H1').activate().mergeAcross();
    newSheet.getRange('A2:H2').activate().mergeAcross();
    newSheet.getRange('A3:H3').activate().mergeAcross();
    newSheet.getRange('A4:H4').activate().mergeAcross();
    newSheet.getRange('A5:H5').activate().mergeAcross();
    newSheet.getRange('A6:H6').activate().mergeAcross();
    newSheet.getRange('A1:H1').activate();
    newSheet.getCurrentCell().setValue(`Information for ${site.site_name} site attendance at ${site.date.toTimeString()} on ${site.date.toLocaleDateString()}`);
    newSheet.getRange('A2:H2').activate();
    newSheet.getCurrentCell().setValue('Fields not persent should be assumed 0.');
    newSheet.getRange('A3:H3').activate();
    newSheet.getCurrentCell().setValue(`Data submitted by ${site["compiled_by"]}`);
    newSheet.getRange('A4:H4').activate();
    newSheet.getCurrentCell().setValue(`Interns present: ${site["present"]}`);
    newSheet.getRange('A5:H5').activate();
    newSheet.getCurrentCell().setValue(`Data leads: ${site["data_leads"]}`);
    newSheet.getRange('A6:H6').activate();
    newSheet.getCurrentCell().setValue(`On-Site lead: ${site["on_site_leads"]}`);
    newSheet.getRange('A8').activate();
    newSheet.getRange('A8:H8').activate().mergeAcross()
    newSheet.getActiveRangeList().setFontStyle('italic')
    newSheet.getCurrentCell().setValue('Combined Totals')
    newSheet.getRange('A16:H16').activate().mergeAcross()
    newSheet.getActiveRangeList().setFontStyle('italic')
    newSheet.getCurrentCell().setValue('Combined Totals (without interns)')
    newSheet.getRange('A24:H24').activate().mergeAcross()
    newSheet.getActiveRangeList().setFontStyle('italic')
    newSheet.getCurrentCell().setValue('Adults')
    newSheet.getRange('A32:H32').activate().mergeAcross()
    newSheet.getActiveRangeList().setFontStyle('italic')
    newSheet.getCurrentCell().setValue('Children')
    newSheet.getRange('A40:H40').activate().mergeAcross()
    newSheet.getActiveRangeList().setFontStyle('italic')
    newSheet.getCurrentCell().setValue('Interns')
    newSheet.getRange('A41').activate();

    /* generate "Age" demographic information */
    var to_insert = get_sums(filtration(site.attached_records, true, true, true), "age", "Age");
    newSheet.getRange("A9:B15").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, true, true), "age", "Age");
    newSheet.getRange("A17:B23").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, true, false), "age", "Age");
    newSheet.getRange("A25:B31").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, false, true), "age", "Age");
    newSheet.getRange("A33:B39").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, true, false, false), "age", "Age");
    newSheet.getRange("A41:B47").setValues(to_insert);

    /* generate "Gender" demographic information */
    to_insert = get_sums(filtration(site.attached_records, true, true, true), "gender", "Gender");
    newSheet.getRange("D9:E15").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, true, true), "gender", "Gender");
    newSheet.getRange("D17:E23").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, true, false), "gender", "Gender");
    newSheet.getRange("D25:E31").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, false, true), "gender", "Gender");
    newSheet.getRange("D33:E39").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, true, false, false), "gender", "Gender");
    newSheet.getRange("D41:E47").setValues(to_insert);

    /* same thing for "Ethnicity" and "Race" */
    to_insert = get_sums(filtration(site.attached_records, true, true, true), "ethnicity", "Ethnicity");
    newSheet.getRange("G9:H15").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, true, true), "ethnicity", "Ethnicity");
    newSheet.getRange("G17:H23").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, true, false), "ethnicity", "Ethnicity");
    newSheet.getRange("G25:H31").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, false, true), "ethnicity", "Ethnicity");
    newSheet.getRange("G33:H39").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, true, false, false), "ethnicity", "Ethnicity");
    newSheet.getRange("G41:H47").setValues(to_insert);

    to_insert = get_sums(filtration(site.attached_records, true, true, true), "race", "Race");
    newSheet.getRange("J9:K15").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, true, true), "race", "Race");
    newSheet.getRange("J17:K23").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, true, false), "race", "Race");
    newSheet.getRange("J25:K31").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, false, false, true), "race", "Race");
    newSheet.getRange("J33:K39").setValues(to_insert);
    to_insert = get_sums(filtration(site.attached_records, true, false, false), "race", "Race");
    newSheet.getRange("J41:K47").setValues(to_insert);
  }

  create_final_report();
}

function get_total_of_attribute(data, attribute, field) {
  var count = 0;
  for(datum of data) {
    if(datum[field] === attribute) {
      count++;
    }
  }

  return count
}

function create_final_report() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.insertSheet();
  try {
    spreadsheet.getActiveSheet().setName('>FINAL REPORT');
  } catch {};
  spreadsheet.getRange('A1:F3').activate()
  .merge();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center')
  .setFontSize(18)
  .setFontSize(18);
  var total_served = filtration(get_demographic_records_for_site(null), false, false, true).length
  spreadsheet.getCurrentCell().setValue(`TOTAL CHILDREN SERVED: ${total_served}`);
  spreadsheet.getActiveRangeList().setVerticalAlignment('middle');
  spreadsheet.getRange('A4:B4').activate()
  .mergeAcross();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('By gender')
  .setTextStyle(0, 9, SpreadsheetApp.newTextStyle()
  .setItalic(true)
  .build())
  .build());
  spreadsheet.getRange('A5').activate();
  spreadsheet.getCurrentCell().setValue('Male');
  spreadsheet.getRange('A6').activate();
  spreadsheet.getCurrentCell().setValue('Female');
  spreadsheet.getRange('A7').activate();
  spreadsheet.getCurrentCell().setValue('Unknown');
  spreadsheet.getRange('A9:B9').activate()
  .mergeAcross();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('By ethnicity')
  .setTextStyle(0, 12, SpreadsheetApp.newTextStyle()
  .setItalic(true)
  .build())
  .build());
  spreadsheet.getRange('A10').activate();
  spreadsheet.getCurrentCell().setValue('Hispanic');
  spreadsheet.getRange('A11').activate();
  spreadsheet.getCurrentCell().setValue('Not Hispanic');
  spreadsheet.getRange('A13:B13').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getActiveRange().mergeAcross();
  spreadsheet.getActiveRangeList().setFontStyle('italic');
  spreadsheet.getCurrentCell().setValue('By race');
  spreadsheet.getRange('A14').activate();
  spreadsheet.getCurrentCell().setValue('American Indian/Alaskian Native');
  spreadsheet.getRange('A15').activate();
  spreadsheet.getCurrentCell().setValue('Asian');
  spreadsheet.getRange('A16').activate();
  spreadsheet.getCurrentCell().setValue('Black/African American');
  spreadsheet.getRange('A17').activate();
  spreadsheet.getCurrentCell().setValue('Native Hawaiian/Other Pacific Islander');
  spreadsheet.getRange('A18').activate();
  spreadsheet.getCurrentCell().setValue('White');
  spreadsheet.getRange('A19').activate();
  spreadsheet.getCurrentCell().setValue('More than one race');
  spreadsheet.getRange('A20').activate();
  spreadsheet.getCurrentCell().setValue('Unknown');
  spreadsheet.getRange('A21').activate();
  spreadsheet.getActiveSheet().setColumnWidth(1, 243);
  spreadsheet.getRangeList(['A13:B13', 'A9:B9', 'A4:B4']).activate()
  .setBackground('#cccccc');

  var data = filtration(get_demographic_records_for_site(null), false, false, true);
  spreadsheet.getRange("B5").setValue(
    get_total_of_attribute(data, "Male", "gender")
  )
  spreadsheet.getRange("B6").setValue(
    get_total_of_attribute(data, "Female", "gender")
  )
  spreadsheet.getRange("B7").setValue(
    get_total_of_attribute(data, "Unknown", "gender")
  )
  spreadsheet.getRange("B10").setValue(
    get_total_of_attribute(data, "Hispanic", "ethnicity")
  )
  spreadsheet.getRange("B11").setValue(
    get_total_of_attribute(data, "Not Hispanic", "ethnicity")
  )
  spreadsheet.getRange("B14").setValue(
    get_total_of_attribute(data, "American Indian/Alaskian Native", "race")
  )
  spreadsheet.getRange("B15").setValue(
    get_total_of_attribute(data, "Asian", "race")
  )
  spreadsheet.getRange("B16").setValue(
    get_total_of_attribute(data, "Black/African American", "race")
  )
  spreadsheet.getRange("B17").setValue(
    get_total_of_attribute(data, "Native Hawaiian/Other Pacific Islander", "race")
  )
  spreadsheet.getRange("B18").setValue(
    get_total_of_attribute(data, "White", "race")
  )
  spreadsheet.getRange("B19").setValue(
    get_total_of_attribute(data, "More than one race", "race")
  )
  spreadsheet.getRange("B20").setValue(
    get_total_of_attribute(data, "Unknown", "race")
  )

  spreadsheet.getRange('A14:B20').activate();
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A14:B20'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setPosition(10, 1, 238, 21)
  .build();
  sheet.insertChart(chart);
  var charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A14:B20'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setPosition(4, 3, 98, 1)
  .build();
  sheet.insertChart(chart);
  spreadsheet.getRange('A5:B7').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B7'));
  chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A5:B7'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'TOTAL SERVED: ${xx}')
  .setPosition(10, 1, 238, 21)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A5:B7'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'TOTAL SERVED: ${xx}')
  .setPosition(21, 3, 98, 15)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A5:B7'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Gender breakdown')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('pieSliceTextStyle.color', '#000000')
  .setOption('titleTextStyle.color', '#757575')
  .setPosition(21, 3, 98, 15)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[0];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A14:B20'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Race breakdown')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('pieSliceTextStyle.color', '#000000')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setPosition(4, 3, 98, 1)
  .build();
  sheet.insertChart(chart);
  spreadsheet.getRange('A10:B11').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B11'));
  chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A10:B11'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setPosition(10, 1, 238, 21)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A10:B11'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setPosition(39, 3, 98, 14)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A10:B11'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setPosition(39, 3, 98, 8)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('A10:B11'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', 'Ethnicity breakdown')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('pieSliceTextStyle.color', '#000000')
  .setOption('titleTextStyle.color', '#757575')
  .setPosition(39, 3, 98, 8)
  .build();
  sheet.insertChart(chart);
  spreadsheet.getRange('A41').activate();
};
