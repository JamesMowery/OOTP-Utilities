/**
 * Adds "OOTP" to the menu, and provides access to various functions
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OOTP')
    .addItem('Clean Salaries', 'cleanSalaries')
    .addItem('Create Salary Totals', 'addSalaries')
    .addSeparator()
    .addItem('Create Budget Estimates', 'addBudgets')
    .addSeparator()
    .addItem('Show Options', 'showSidebar')
    .addToUi();
}

/**
 * Prompts the user to enter budgets and returns those values
 */
function getBudgets() {
  var ui = SpreadsheetApp.getUi();

  var result;
  // var button, text

  var budgets = {
    'last': 0,
    'current': 0,
    'next': 0,
    'two': 0,
  };

  result = ui.prompt(
      'Let\'s set up your budgets!',
      'What was your previous year\'s budget?',
      ui.ButtonSet.OK_CANCEL);
  budgets.last = Number(result.getResponseText());

  result = ui.prompt(
      'This Year',
      'What is this year\'s projected budget?',
      ui.ButtonSet.OK_CANCEL);
  budgets.current = Number(result.getResponseText());

  result = ui.prompt(
      'Next Year',
      'What is your budget projected to be next year?',
      ui.ButtonSet.OK_CANCEL);
  budgets.next = Number(result.getResponseText());

  result = ui.prompt(
      'Two Years Ahead',
      'What is your budget projected to be in two years?',
      ui.ButtonSet.OK_CANCEL);
  budgets.two = Number(result.getResponseText());

  /*
  // Process the user's response.
  button = result.getSelectedButton();
  text = result.getResponseText();
  if (button == ui.Button.OK) {
    // Feedback, if needed
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('Your budgets were not set.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('Your budgets were not set.');
  }
  */

  return budgets;
}

/**
 * Generates a projection of the future team budget
 * by using Google Sheet's TREND formula.
 */
function addBudgets() {

  var sheet = SpreadsheetApp.getActiveSheet();

  var budgetTerm = "BUDGET";
  var selectedTermBudget = false;
  var currentRow = 0;
  var currentCol = 0;
  var numberFormat = "$#,##0_)";

  // If selected cell contains budgetTerm,
  // insert TREND formula to calculate future budgets
  if (sheet.getActiveCell().getValue() == budgetTerm) {
    // TODO: If selected cell contains budgetTerm
    // expand formulas to the right
    var data = sheet.getDataRange().getValues();
  }
  else {
    var data = sheet.getDataRange().getValues();
  }

  // Search for the term denoting the budget row
  for (var i in data) {
    for (var j in data[i]) {

      // If the budgetTerm is found, store the cell
      if (String(data[i][j]).search(budgetTerm) !== -1) {
        currentRow = Number(i) + 1;
        currentCol = Number(j);

      }
    }
  }

  // If the budgetTerm is not found, find the end of the sheet,
  // and insert the budgetTerm in the first column
  if (currentRow === 0 && currentCol === 0) {
    currentRow = Number(sheet.getDataRange().getHeight()) + 1;
    sheet.getRange(currentRow, 1).setValue(budgetTerm);
  }

  // Retrieve the budgets from the user
  var budgets = getBudgets();

  // Set the budget values in the cells
  sheet.getRange(currentRow, 2).setValue(budgets.last)
    .setNumberFormat(numberFormat);
  sheet.getRange(currentRow, 3).setValue(budgets.current)
    .setNumberFormat(numberFormat);
  sheet.getRange(currentRow, 4).setValue(budgets.next)
    .setNumberFormat(numberFormat);
  sheet.getRange(currentRow, 5).setValue(budgets.two)
    .setNumberFormat(numberFormat);

  // Set the TREND formula for the remaining columns
  sheet.getRange(currentRow, 6)
                 .setValue(Utilities
                   .formatString('=TREND(B%s:E%s, B1:E1, F1:K1)',
                                 currentRow, currentRow))
                 .setNumberFormat(numberFormat);

}

/**
 * A function that detects the totals row, and adds
 * SUM formulas for each of the columns.
 */
function addSalaries() {
  var sheet = SpreadsheetApp.getActiveSheet();

  var totalTerm = "TOTAL";
  var selectedTermTotal = false;
  var cell = [];
  var currentRow = 0;
  var currentCol = 0;
  var numberFormat = "$#,##0_)";

  // If selected cell contains totalTerm, expand SUM formulas,
  // otherwise collect all data and search for totalTerm
  if (sheet.getActiveCell().getValue() == totalTerm) {
    // TODO: If selected cell contains totalTerm
    // expand formulas to the right
  }
  else {
    var data = sheet.getDataRange().getValues();
  }

  // Search for the term denoting the totals row
  for (var i in data) {
    for (var j in data[i]) {
      if (String(data[i][j]).search(totalTerm) !== -1) {

        // When found, store the row and column values
        currentRow = Number(i);
        currentCol = Number(j);
      }
    }
  }

  // Inserts the SUM formulas in the row that represents totals
  sheet.getRange(currentRow + 1, currentCol + 2,
                 1, sheet.getDataRange().getWidth() - 1)
                 .setValue(Utilities.formatString('=SUM(B2:B%s)', currentRow))
                 .setNumberFormat(numberFormat);
}

/**
 * A function that cleans and reformats OOTP salary data,
 * thus making it usable within a spreadsheet environment.
 */
function cleanSalaries() {
  var sheet = SpreadsheetApp.getActiveSheet();

  var result = [];
  var cellText = "";
  var leadingSlice = 0;
  var leadingNumber = "";
  var numberFormat = "$#,##0_)";

  var rangeSelectorEnabled = false;

  // TODO: Detect if custom range selection is enabled via dialog box

  if (rangeSelectorEnabled === true) {
    // var data = SpreadsheetApp.getActiveSheet().getActiveRange().getValues();
  }
  else {
    var data = sheet.getRange(2, 2, sheet.getDataRange().getHeight() - 2,
                              sheet.getDataRange().getWidth() - 2).getValues();

    // Freeze the first row and column for pretty formatting
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }

  // Search through each cell to locate data that needs modified,
  // and then apply those modifications.
  for (var i = 0; i < data.length; i++) {

    // Push an empty array to enable 2D array
    result.push([]);

    for (var j = 0; j < data[i].length; j++) {

      // Convert data to a string for modification
      cellText = String(data[i][j]);

      // Remove symbols, commas, and periods
      cellText = cellText.replace(/[,./$]/g, "");

      // Replace millions (m) with zeros
      if (cellText.search(/[^A-Za-z]m(?=\()?/g) !== -1) {

        // Locate the "m", and get the leading number
        leadingSlice = cellText.search(/[^A-Za-z]m(?=\()?/g);
        leadingNumber = cellText.slice(leadingSlice, leadingSlice + 1);

        // If the leading number isn't a 0, include the leading number,
        // and then add the appropriate number of zeros.
        if (leadingNumber != "0") {
          // Add the leading number
          cellText = cellText.replace(/[^A-Za-z]m(?=\()?/g,
                                      leadingNumber + "00000");
        }
        else {
          // No leading number, just add zeros
          cellText = cellText.replace(/[^A-Za-z]m(?=\()?/g, "000000");
        }

      } // if millions

      // Replace thousands (k) with zeros
      if (cellText.search(/[^A-Za-z]k(?=\()?/g) !== -1) {

        // Locate "k", and get the leading number
        leadingSlice = cellText.search(/[^A-Za-z]k(?=\()?/g);
        leadingNumber = cellText.slice(leadingSlice, leadingSlice + 1);

        /**
         * If the leading number isn't a 0, include the leading number,
         * and then add the appropriate number of zeros.
         */
        if (leadingNumber != "0") {
          // Add the leading number
          cellText = cellText.replace(/[^A-Za-z]k(?=\()?/g,
                                      leadingNumber + "000");
        }
        else {
          // No leading number, just add zeros
          cellText = cellText.replace(/[^A-Za-z]k(?=\()?/g, "0000");
        }

      } // if thousands

      // Remove parenthetical information
      if (cellText.search(/(\(.+\))$/g) !== -1) {
          cellText = cellText.replace(/(\(.+\))$/g, "");
      }

      // Assign the modifications to the array
      result[i][j] = cellText;

    } // Inner for loop
  } // Outer for loop

  // If using a selected range, apply results to that data range,
  // otherwise apply the results to the default range.
  if (rangeSelectorEnabled === true) {
    // sheet.getActiveRange().setValues(result).setNumberFormat(numberFormat);
  }
  else {
    sheet.getRange(2, 2, sheet.getDataRange().getHeight() - 2,
                   sheet.getDataRange().getWidth() - 2).setValues(result);
  }

  // Set the formatting of the numbers
  sheet.getRange(2, 2, sheet.getDataRange().getHeight() - 2,
                 10).setNumberFormat(numberFormat);

}
