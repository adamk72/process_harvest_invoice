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

function GetConfigInformation(activeSS) {
  let dateFormat = 'yyyy-MM-dd';
  let configSS = activeSS.getSheetByName('Config');

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
    startdate,
    endDate,
  };
}

function ProcessInvoice() {
  let dateFormat = 'yyyy-MM-dd';

  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let activeSS = SpreadsheetApp.setActiveSheet(
    spreadsheet.getSheetByName('Hours'),
    true
  );
  let cfg = GetConfigInformation(activeSS);
  let userTimeZone = spreadsheet.getSpreadsheetTimeZone();

  let dataRangeHours = activeSS.getDataRange();
  let dataHours = dataRangeHours.getValues();

  let billSum = 0;
  let hoursTasksProjects = [];

  let startDataCollecting = false;

  let projectTaskDict = {};
  let projectId = null;

  for (let i = 0; i < dataHours.length; i++) {
    if (typeof dataHours[i][0] == 'string') {
      continue;
    }
    let dataDate = Utilities.formatDate(
      new Date(dataHours[i][0]),
      userTimeZone,
      dateFormat
    );
    let taskCell = null;
    let hoursCell = null;
    let projectCell = null;
    let j = i + 1;
    if (dataDate == startDate && !startDataCollecting) {
      startDataCollecting = true;
      taskCell = dataRangeHours.getCell(j, 5);
      hoursCell = dataRangeHours.getCell(j, 7);
      projectCell = dataRangeHours.getCell(j, 3);
      hoursTasksProjects.push({
        project: projectCell.getValue(),
        task: taskCell.getValue(),
        hours: hoursCell.getValue(),
        date: dataDate,
      });

      let taskName = taskCell.getValue();
      let currentHours = projectTaskDict.projectTaskId.taskName;
      projectTaskDict[projectId][taskName] = {
        hours: currentHours + hoursCell.getValue(),
      };

      continue; // Got the first date, so continue in order to incrementing
    }
    if (dataDate > endDate) {
      break;
    } // a bigger date, so we're done.
    if (dataDate >= startDate && startDataCollecting) {
      taskCell = dataRangeHours.getCell(j, 5);
      hoursCell = dataRangeHours.getCell(j, 7);
      projectCell = dataRangeHours.getCell(j, 3);

      hoursTasksProjects.push({
        project: projectCell.getValue(),
        task: taskCell.getValue(),
        hours: hoursCell.getValue(),
        date: dataDate,
      });
      let taskName = taskCell.getValue();
      let currentHours = projectTaskDict.projectTaskId.taskName;
      projectTaskDict[projectId][taskName] = {
        hours: currentHours + hoursCell.getValue(),
      };
    }
  }

  let projObj = {};
  hoursTasksProjects.forEach((obj) => {
    if (!Array.isArray(projObj[obj.project])) {
      projObj[obj.project] = [];
    }
    projObj[obj.project].push({
      task: obj.task,
      hours: obj.hours,
      date: obj.date,
    });
  });

  // Now let's start populating the actual invoice sheet

  spreadsheet.setActiveSheet(
    spreadsheet.getSheetByName('Invoice Master'),
    true
  );
  let newInvoiceName =
    'Invoice for ' +
    Utilities.formatDate(new Date(startDate), 'GMT+1', 'MM/dd/YY') +
    '-' +
    Utilities.formatDate(new Date(endDate), 'GMT+1', 'MM/dd/YY');

  if (spreadsheet.getSheetByName(newInvoiceName) == null) {
    spreadsheet.duplicateActiveSheet();
    spreadsheet
      .getSheetByName('Copy of Invoice Master')
      .setName(newInvoiceName);
  }

  activeSS = spreadsheet.setActiveSheet(
    spreadsheet.getSheetByName(newInvoiceName),
    true
  );

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

  for (const proj in projObj) {
    // collate the hours
    let taskObj = {};
    for (let k = 0; k < projObj[proj].length; k++) {
      let task = projObj[proj][k].task,
        date = projObj[proj][k].date,
        hours = projObj[proj][k].hours;

      if (taskObj[task] == null) taskObj[task] = {};
      if (taskObj[task] != null && taskObj[task][date] == null)
        taskObj[task][date] = 0;
      taskObj[task][date] += hours;
    }

    let cell = currentRowRange.getCell(1, 2);
    cell.setValue(proj);
    cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    for (const task in taskObj) {
      let perDateCount = Object.size(taskObj[task]);
      cell = currentRowRange.getCell(1, 3);
      cell.setValue(task);
      cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      for (const [taskDate, taskHours] of Object.entries(taskObj[task])) {
        cell = currentRowRange.getCell(1, 4); // date
        cell.setValue(taskDate);

        cell = currentRowRange.getCell(1, 5); // hours
        cell.setValue(taskHours);

        cell = currentRowRange.getCell(1, 6); // rate
        cell.setValue(hourlyRate);

        cell = currentRowRange.getCell(1, 7); // total
        cell.setFormula('=E' + currentRow + '*F' + currentRow);

        if ((currentRow - 1) % 2) {
          activeSS.getRange(currentRow, 2, 1, 6).setBackground(altRowColor);
        } else {
          activeSS.getRange(currentRow, 2, 1, 6).setBackground('#ffffff');
        }

        if (perDateCount-- > 0) {
          activeSS.insertRowAfter(currentRow);
          currentRow++;
          currentRowRange = activeSS.getRange(
            'A' + currentRow + ':H' + currentRow
          ); // TODO: make this more seamless; currentRow and currentRowRange risk messing up each other if I make other code changes.
        }
      }
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
    .setFontColor(defocusedTextColor)
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
