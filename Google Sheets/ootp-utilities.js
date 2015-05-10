/**
 * Adds "OOTP" to the menu, and provides access to various functions
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OOTP')
    .addItem('Color Cells', 'colorCells')
    .addItem('Format Salaries', 'cleanSalaries')
    .addItem('Compute Salary Totals', 'addSalaries')
    .addSeparator()
    .addItem('Create Budget Estimates', 'addBudgets')
    .addItem('Create Remaining Budget', 'remainingBudget')
    .addSeparator()
    .addItem('Show Options', 'showSidebar')
    .addToUi();
}

/**
 * Locates cells that have contract data
 */
function findContractCells(sheet, data) {
  var coloredCells = {
    'playerOptionCells': [],
    'teamOptionCells': [],
    'vestingOptionCells': [],
    'autoContractCells': [],
    'arbitrationCells': [],
    'minorContractCells': []
  };

  var i, j;

  // Search for contract specifications
  for (i in data) {
    for (j in data[i]) {

      // If the contract is a player option
      if (String(data[i][j]).search(/\(P\)$/g) !== -1) {
        coloredCells.playerOptionCells.push([Number(i) + 1, Number(j) + 1]);
      }

      // If the contract is a team option
      if (String(data[i][j]).search(/\(T\)$/g) !== -1) {
        coloredCells.teamOptionCells.push([Number(i) + 1, Number(j) + 1]);
      }

      // If the contract is a vesting option
      if (String(data[i][j]).search(/\(V\)$/g) !== -1) {
        coloredCells.vestingOptionCells.push([Number(i) + 1, Number(j) + 1]);
      }

      // If the contract is a auto contract
      if (String(data[i][j]).search(/\(auto\)$/g) !== -1) {
        coloredCells.autoContractCells.push([Number(i) + 1, Number(j) + 1]);
      }

      // If the contract is a minor league contract
      if (String(data[i][j]).search(/\(MiLC\)$/g) !== -1) {
        coloredCells.minorContractCells.push([Number(i) + 1, Number(j) + 1]);
      }

      // If the contract is possibly for arbitration
      if (String(data[i][j]).search(/\(A.?\)$/g) !== -1) {
        coloredCells.arbitrationCells.push([Number(i) + 1, Number(j) + 1]);
      }

    } // for inner loop
  } // for outer loop

  return coloredCells;
}

/**
 * Applies color to specific cells based on contract status
 */
function colorCells() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // Locate the cells with contract data
  var cells = findContractCells(sheet, data);

  var playerOptionColor;
  var teamOptionColor;
  var vestingOptionColor;
  var autoContractColor;
  var arbitrationColor;
  var minorContractColor;

  var i;

  // Color player option cells
  for (i in cells.playerOptionCells) {
    sheet.getRange(cells.playerOptionCells[i][0],
      cells.playerOptionCells[i][1]).setBackgroundRGB(250, 135, 135);
  }

  // Color team option cells
  for (i in cells.teamOptionCells) {
    sheet.getRange(cells.teamOptionCells[i][0],
      cells.teamOptionCells[i][1]).setBackgroundRGB(135, 176, 230);
  }

  // Color vesting option cells
  for (i in cells.vestingOptionCells) {
    sheet.getRange(cells.vestingOptionCells[i][0],
      cells.vestingOptionCells[i][1]).setBackgroundRGB(230, 135, 206);
  }

  // Color auto contract cells
  for (i in cells.autoContractCells) {
    sheet.getRange(cells.autoContractCells[i][0],
      cells.autoContractCells[i][1]).setBackgroundRGB(135, 230, 179);
  }

  // Color arbitration contract cells
  for (i in cells.arbitrationCells) {
    sheet.getRange(cells.arbitrationCells[i][0],
      cells.arbitrationCells[i][1]).setBackgroundRGB(230, 230, 135);
  }

  // Color minor league contract cells
  for (i in cells.minorContractCells) {
    sheet.getRange(cells.minorContractCells[i][0],
      cells.minorContractCells[i][1]).setBackgroundRGB(135, 230, 228);
  }

  // TODO: Add option to remove all coloring
}

/**
 * Displays the remaining budget by subtracting budget from salary
 */
function remainingBudget() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();

  var hasBudget = false;
  var budgetTerm = "BUDGET";
  var i, j;
  var response = "";
  var lastRow;
  var currentRow;
  var numberFormat = "$#,##0_)";

  // Check to see if budget estimates have been rendered
  for (i in data) {
    for (j in data[i]) {
      // If the budgetTerm is found, continue
      if (String(data[i][j]).search(budgetTerm) !== -1) {
        currentRow = Number(i) + 1;
        hasBudget = true;
      }
    }
  }

  // If the budgetTerm isn't found, ask the user to create one or cancel
  if (hasBudget === false) {

    response = ui.alert("No budget has been created.",
                            "Would you like to create a budget now? \
                        \(If not, this operation will cancel.\)",
                            ui.ButtonSet.YES_NO);
    // Create the budgets, or stop the operation, depending on the response
    if (String(response) ===  "YES") {
      addBudgets();
    }
    else if (String(response) === "NO") {
      return null;
    }

    // Select the last row
    currentRow = Number(sheet.getDataRange().getHeight()) + 1;
  }

  sheet.getRange(currentRow + 1, 1).setValue("REMAINING");
  sheet.getRange(currentRow + 1, 2, 1,
                 Number(sheet.getDataRange().getWidth()) - 1)
                 .setValue(Utilities.formatString('=MINUS(B%s, B%s)', currentRow, currentRow - 1))
                 .setNumberFormat(numberFormat);
}

/**
 * Prompts the user to enter budgets and returns those values
 */
function getBudgets() {
  var ui = SpreadsheetApp.getUi();

  var result = null;
  var budgets = {
      'last': 0,
      'current': 0,
      'next': 0,
      'two': 0
    };

  // Prompt the user for last year's budget
  result = ui.prompt(
    'Let\'s set up your budgets!',
    'What was your previous year\'s budget?',
    ui.ButtonSet.OK_CANCEL
  );
  budgets.last = Number(result.getResponseText());

  // Prompt the user for this year's budget
  result = ui.prompt(
    'This Year',
    'What is this year\'s projected budget?',
    ui.ButtonSet.OK_CANCEL
  );
  budgets.current = Number(result.getResponseText());

  // Prompt the user for next year's projected budget
  result = ui.prompt(
    'Next Year',
    'What is your budget projected to be next year?',
    ui.ButtonSet.OK_CANCEL
  );
  budgets.next = Number(result.getResponseText());

  // Prompt the user for the projected budget in two years
  result = ui.prompt(
    'Two Years Ahead',
    'What is your budget projected to be in two years?',
    ui.ButtonSet.OK_CANCEL
  );
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
  var data = null;
  var i = 0, j = 0;
  var budgets = null;

  // If selected cell contains budgetTerm,
  // insert TREND formula to calculate future budgets
  if (sheet.getActiveCell().getValue() === budgetTerm) {
    // TODO: If selected cell contains budgetTerm
    // expand formulas to the right
    data = sheet.getDataRange().getValues();
  }
  else {
    data = sheet.getDataRange().getValues();
  }

  // Search for the term denoting the budget row
  for (i in data) {
    for (j in data[i]) {

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
  budgets = getBudgets();

  // Set the budget values in the cells
  sheet.getRange(currentRow, 2).setValue(budgets.last);
  sheet.getRange(currentRow, 3).setValue(budgets.current);
  sheet.getRange(currentRow, 4).setValue(budgets.next);
  sheet.getRange(currentRow, 5).setValue(budgets.two);

  // Set the TREND formula for the remaining columns
  sheet.getRange(currentRow, 6)
    .setValue(
      Utilities.formatString(
        '=TREND(B%s:E%s, B1:E1, F1:K1)',
        currentRow, currentRow
      )
    );

  // Set the number formats for the column
  sheet.getRange(currentRow, 2, 1,
                 sheet.getDataRange().getWidth() - 1
                ).setNumberFormat(numberFormat);
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
  var data = null;
  var i, j;

  // If selected cell contains totalTerm, expand SUM formulas,
  // otherwise collect all data and search for totalTerm
  if (sheet.getActiveCell().getValue() === totalTerm) {
    // TODO: If selected cell contains totalTerm
    // expand formulas to the right
  }
  else {
    data = sheet.getDataRange().getValues();
  }

  // Search for the term denoting the totals row
  for (i in data) {
    for (j in data[i]) {
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
  var data = null;
  var i, j;

  // TODO: Detect if custom range selection is enabled via dialog box

  if (rangeSelectorEnabled === true) {
    // data = SpreadsheetApp.getActiveSheet().getActiveRange().getValues();
  }
  else {
    data = sheet.getRange(2, 2, sheet.getDataRange().getHeight() - 2,
                          sheet.getDataRange().getWidth() - 2)
                          .getValues();

    // Freeze the first row and column for pretty formatting
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }

  // Search through each cell to locate data that needs modified,
  // and then apply those modifications.
  for (i = 0; i < data.length; i++) {

    // Push an empty array to enable 2D array
    result.push([]);

    for (j = 0; j < data[i].length; j++) {

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
