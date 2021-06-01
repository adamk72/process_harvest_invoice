/**
 * @author Adam Kecskes
 * @note Work in progress; use at your own risk. Also, this is hardly an efficient method for pulling data.
 * @version 0.0.1
 **/

Object.size = function (obj) {
  let size = 0,
    key;
  for (key in obj) {
    if (obj.hasOwnProperty(key)) size++;
  }
  return size;
};

function GetConfigInformation(ss) {
  let dateFormat = 'yyyy-MM-dd';
  let configSS = ss.getSheetByName('Config');
  let userTimeZone = ss.getSpreadsheetTimeZone();

  let hourlyRate = configSS
    .getRange('A:A')
    .createTextFinder('Hourly Rate')
    .matchCase(true)
    .findNext()
    .offset(0, 1)
    .getValue();
  let altRowColor = configSS
    .getRange('A:A')
    .createTextFinder('Alt Row Hex Color')
    .matchCase(true)
    .findNext()
    .offset(0, 1)
    .getValue();
  let defocusedTextColor = configSS
    .getRange('A:A')
    .createTextFinder('Defocused Hex Color')
    .matchCase(true)
    .findNext()
    .offset(0, 1)
    .getValue();

  let startDate = configSS
    .getRange('A:A')
    .createTextFinder('Start Date')
    .matchCase(true)
    .findNext()
    .offset(0, 1)
    .getValue();
  let endDate = configSS
    .getRange('A:A')
    .createTextFinder('End Date')
    .matchCase(true)
    .findNext()
    .offset(0, 1)
    .getValue();
  startDate = Utilities.formatDate(startDate, userTimeZone, dateFormat);
  endDate = Utilities.formatDate(endDate, userTimeZone, dateFormat);

  return {
    rate: hourlyRate,
    altRowColor,
    defocusedTextColor,
    startDate,
    endDate,
  };
}

function ProcessInvoice() {
  let dateFormat = 'yyyy-MM-dd';

  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let activeSS = SpreadsheetApp.setActiveSheet(
    sheet.getSheetByName('Hours'),
    true
  );
  let cfg = GetConfigInformation(sheet);
  let userTimeZone = sheet.getSpreadsheetTimeZone();

  let dateRange = activeSS.getDataRange();
  let dataValues = dateRange.getValues();

  let billSum = 0;
  let hoursTasksProjects = [];

  let startDataCollecting = false;

  let projectTaskDict = {};

  for (let dateIndex = 0; dateIndex < dataValues.length; dateIndex++) { 

    if (typeof dataValues[dateIndex][0] == 'string') {
      continue;
    }
    let dataDate = Utilities.formatDate(
      new Date(dataValues[dateIndex][0]),
      userTimeZone,
      dateFormat
    );

    if (dataDate > cfg.endDate) {
      break;
    } // a bigger date, so we're done.

    if (dataDate >= cfg.startDate && !startDataCollecting) {
      // console.log(dataDate, cfg.startDate);
      startDataCollecting = true;
    }

    if (startDataCollecting) {
      let projectName = dateRange.getCell(dateIndex + 1, 3).getValue();
      let taskName = dateRange.getCell(dateIndex + 1, 5).getValue();
      let hours = dateRange.getCell(dateIndex + 1, 7).getValue();

      projectTaskDict[projectName] = projectTaskDict[projectName] || {};
      projectTaskDict[projectName][taskName] =
        projectTaskDict[projectName][taskName] || 0;

      let currentHours = Number(projectTaskDict[projectName][taskName]) + hours;
      projectTaskDict[projectName][taskName] = Number(currentHours).toFixed(2);
    }
  }

  // Now let's start populating the actual invoice sheet
  sheet.setActiveSheet(sheet.getSheetByName('Invoice Master'), true);
  let newInvoiceName =
    'Invoice for ' +
    Utilities.formatDate(new Date(cfg.startDate), 'GMT+1', 'MM/dd/YY') +
    '-' +
    Utilities.formatDate(new Date(cfg.endDate), 'GMT+1', 'MM/dd/YY');

  if (sheet.getSheetByName(newInvoiceName) == null) {
    sheet.duplicateActiveSheet();
    sheet.getSheetByName('Copy of Invoice Master').setName(newInvoiceName);
  }

  activeSS = sheet.setActiveSheet(sheet.getSheetByName(newInvoiceName), true);

  // Delete the current list
  // ProjectSTART is a unique word; "START" is the same color as cell background so it will not be seen. This is how I'll find where the data to delete is.
  let startCell = activeSS
    .getRange('B:B')
    .createTextFinder('ProjectSTART')
    .matchCase(true)
    .findNext();
  let startRow = startCell.getRow() + 1; // want to focus on the row where change will actually happen so the header remains untouched.

  let endCell = activeSS
    .getRange('B:B')
    .createTextFinder('Notes:')
    .matchCase(true)
    .findNext();
  let lastRow = endCell.getRow();

  // Count current rows so we know how many to delete.
  // Delete the rows, then we're going to add new rows one at a time
  if (lastRow - startRow > 0) activeSS.deleteRows(startRow, lastRow - startRow);

  let currentRow = startRow;
  activeSS.insertRowBefore(currentRow);
  let currentRowRange = activeSS.getRange('A' + currentRow + ':H' + currentRow);

  for (let projectName of Object.keys(projectTaskDict)) {
    let cell = currentRowRange.getCell(1, 2);
    cell.setValue(projectName);
    cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    for (let taskName of Object.keys(projectTaskDict[projectName])) {
      cell = currentRowRange.getCell(1, 3);
      cell.setValue(taskName);
      cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

      cell = currentRowRange.getCell(1, 5); // hours
      cell.setValue(projectTaskDict[projectName][taskName]);

      cell = currentRowRange.getCell(1, 6); // rate
      cell.setValue(cfg.rate);

      cell = currentRowRange.getCell(1, 7); // total
      cell.setFormula('=E' + currentRow + '*F' + currentRow);

      if ((currentRow - 1) % 2) {
        activeSS.getRange(currentRow, 2, 1, 6).setBackground(cfg.altRowColor);
      } else {
        activeSS.getRange(currentRow, 2, 1, 6).setBackground('#ffffff');
      }

      activeSS.insertRowAfter(currentRow);
      currentRow++;
      currentRowRange = activeSS.getRange('A' + currentRow + ':H' + currentRow);
    }
  }

  let lastInvoiceRow = currentRow + 1;
  let numOfRows = currentRow - startRow;

  // Add sum function for total
  activeSS
    .getRange(lastInvoiceRow, 7, 1, 1)
    .setFormula(`=sum(G${startRow}:G${lastInvoiceRow - 1})`);

  // General formatting
  activeSS
    .getRange(startRow, 2, numOfRows, 6)
    .setFontColor(cfg.defocusedTextColor)
    .setFontWeight(null); // all
  activeSS
    .getRange(startRow, 2, numOfRows, 1)
    .setFontColor('#000000')
    .setFontWeight(null); // project
  activeSS
    .getRange(startRow, 3, numOfRows, 1)
    .setFontColor('#000000')
    .setFontWeight(null); // task
  activeSS
    .getRange(startRow, 4, numOfRows, 1)
    .setNumberFormat('M/d/yy')
    .setHorizontalAlignment('left'); // dates
  activeSS.getRange(startRow, 6, numOfRows, 1).setNumberFormat('"$"#,##0.00'); // rate only
}
