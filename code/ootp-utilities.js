/**
 * Adds "OOTP" to the menu, and provides access to various functions
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OOTP')
    .addItem('Format Data & Compute Totals', 'doEverything')
    .addItem('Format Data & Compute Totals (With Color)', 'doEverythingColor')
    .addSeparator()
    .addItem('Format Data Only', 'cleanSalaries')
    .addItem('Format Data Only \(With Color\)', 'cleanSalariesColor')
    .addSeparator()
    .addItem('Compute All Totals \(After Formatting Data\)', 'generateAllTotals')
    .addSeparator()
    .addItem('Compute Salary Totals', 'addSalaries')
    .addItem('Compute Budget Estimates', 'addBudgets')
    .addItem('Compute Remaining Budget', 'remainingBudget')
    .addSeparator()
    .addItem('Add/Update Cell Coloring \(Before Formatting Data\)',
             'colorCells')
    .addItem('Remove Cell Coloring', 'removeColor')
    .addSeparator()
    .addItem('Add/Reset Settings Sheet', 'generateSettingsSheet')
    .addToUi();
}

/**
 * Executes all primary functions after adding color
 */
function doEverythingColor() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  var settingsSheet = spreadsheet.getSheetByName("settings");

  if (settingsSheet == null) {
    generateSettingsSheet();
  }

  sheet.activate();

  colorCells();
  cleanSalaries();
  generateAllTotals();
}

/**
 * Executes all primary functions
 */
function doEverything() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  var settingsSheet = spreadsheet.getSheetByName("settings");

  if (settingsSheet == null) {
    generateSettingsSheet();
  }

  sheet.activate();

  cleanSalaries();
  generateAllTotals();
}

/**
 * Executes all functions related to computing totals
 */
function generateAllTotals() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  var settingsSheet = spreadsheet.getSheetByName("settings");

  if (settingsSheet == null) {
    generateSettingsSheet();
  }

  sheet.activate();

  addSalaries();
  addBudgets();
  remainingBudget();
}

/**
 * Retreives an individual setting
 */
function getSetting(option) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var cell = null;

  var sheet = spreadsheet.getSheetByName("settings");

  if (sheet == null) {
    ui.alert("Settings Sheet Not Found",
      "A settings sheet could not be located. One will now be created.",
      ui.ButtonSet.OK);
      generateSettingsSheet();
      sheet = spreadsheet.getSheetByName("settings");
  }

  // Target the cell of the requested setting
  switch (option) {
    case "freeze":
      cell = sheet.getRange(2, 2, 1, 1);
      break;
    case "player":
      cell = sheet.getRange(5, 2, 1, 1);
      break;
    case "team":
      cell = sheet.getRange(6, 2, 1, 1);
      break;
    case "vesting":
      cell = sheet.getRange(7, 2, 1, 1);
      break;
    case "auto":
      cell = sheet.getRange(8, 2, 1, 1);
      break;
    case "arbitration":
      cell = sheet.getRange(9, 2, 1, 1);
      break;
    case "minor":
      cell = sheet.getRange(10, 2, 1, 1);
      break;
    case "format":
      cell = sheet.getRange(13, 2, 1, 1);
      break;
    case "salary":
      cell = sheet.getRange(16, 2, 1, 1);
      break;
    case "budget":
      cell = sheet.getRange(17, 2, 1, 1);
      break;
    case "remaining":
      cell = sheet.getRange(18, 2, 1, 1);
      break;
    default:
      cell = undefined;
      break;
  }

  if (cell == null || cell == undefined || cell == "") {
    var response = ui.alert("Setting Not Found",
      "Your settings sheet might be broken. Reset it to default?",
      ui.ButtonSet.OK_CANCEL);
    if (response == "OK") {
      generateSettingsSheet();
    }
    else {
      return null;
    }
  }

  // If an appropriate cell was chosen, return it
  if (cell !== undefined || cell !== null || cell !== "") {
    return cell.getValue();
  }
}

/**
 * Fills in a settings sheet with the initial settings
 */
function populateSettingsSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("settings");

  var cell = null;
  var rule = null;

  cell = sheet.getRange(1, 1, 1, 1).setValue("General Options")
         .setFontWeight("Bold");
  cell = sheet.getRange(1, 2, 1, 1).setValue("Value \(Yes/No\)")
         .setFontWeight("Bold");

  cell = sheet.getRange(2, 1, 1, 1).setValue("Freeze First Row/Column");
  cell = sheet.getRange(2, 2, 1, 1).setValue("Yes");
  // Set validation to require a Yes/No answer with dropdown
  rule = SpreadsheetApp.newDataValidation()
         .requireValueInList(['Yes', 'No'], true).build();
  cell.setDataValidation(rule);

  cell = sheet.getRange(4, 1, 1, 1).setValue("Color Options")
         .setFontWeight("Bold");
  cell = sheet.getRange(4, 2, 1, 1).setValue("Color Value \(#XXXXXX\)")
         .setFontWeight("Bold");

  cell = sheet.getRange(5, 1, 1, 1).setValue("Player Option");
  cell = sheet.getRange(5, 2, 1, 1).setValue("#FB8072");

  cell = sheet.getRange(6, 1, 1, 1).setValue("Team Option");
  cell = sheet.getRange(6, 2, 1, 1).setValue("#B7D2FF");

  cell = sheet.getRange(7, 1, 1, 1).setValue("Vesting Option");
  cell = sheet.getRange(7, 2, 1, 1).setValue("#BEBADA");

  cell = sheet.getRange(8, 1, 1, 1).setValue("Auto Contract");
  cell = sheet.getRange(8, 2, 1, 1).setValue("#8DD3C7");

  cell = sheet.getRange(9, 1, 1, 1).setValue("Arbitration");
  cell = sheet.getRange(9, 2, 1, 1).setValue("#FFFFB3");

  cell = sheet.getRange(10, 1, 1, 1).setValue("Minor League");
  cell = sheet.getRange(10, 2, 1, 1).setValue("#ECECEC");

  cell = sheet.getRange(12, 1, 1, 1).setValue("Number Format")
         .setFontWeight("Bold");
  cell = sheet.getRange(12, 2, 1, 1).setValue("For Help, Visit:")
         .setFontWeight("Bold");
  cell = sheet.getRange(12, 3, 1, 1)
         .setValue("https://support.google.com/docs/answer/56470?hl=en");

  cell = sheet.getRange(13, 1, 1, 1).setValue("Format");
  cell = sheet.getRange(13, 2, 1, 1).setValue("$#,##0_)");

  cell = sheet.getRange(15, 1, 1, 1).setValue("Row Text Descriptions")
         .setFontWeight("Bold");
  cell = sheet.getRange(15, 2, 1, 1).setValue("Text")
         .setFontWeight("Bold");

  cell = sheet.getRange(16, 1, 1, 1).setValue("Salary Total");
  cell = sheet.getRange(16, 2, 1, 1).setValue("TOTAL");

  cell = sheet.getRange(17, 1, 1, 1).setValue("Budget Total");
  cell = sheet.getRange(17, 2, 1, 1).setValue("BUDGET");

  cell = sheet.getRange(18, 1, 1, 1).setValue("Remaining Total");
  cell = sheet.getRange(18, 2, 1, 1).setValue("REMAINING");
}

/**
 * (Re)Generates a settings sheet
 */
function generateSettingsSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var result = null;

  // Locate the the "settings" sheet
  var sheet = spreadsheet.getSheetByName("settings");

  // If the sheet exists, ask if the user would like to reset it
  if (sheet !== null) {
    result = ui.alert("Settings sheet already exists",
                      "Would you like to reset your settings sheet?",
                      ui.ButtonSet.YES_NO);
    if (result == "YES") {
      sheet.clear().activate();
      populateSettingsSheet();
    }
    else {
      sheet.activate();
    }
  }
  // If the sheet doesn't exist, create a new sheet called "settings"
  else {
    sheet = spreadsheet.insertSheet().setName("settings").activate();
    populateSettingsSheet();
  }

  return sheet;
}

/**
 * Removes background colors from the spreadsheet
 */
function removeColor() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var totalTerm = getSetting("salary");

  var totalCols = Number(sheet.getDataRange().getWidth()) - 1;
  var lastNumberRow, i, j;

  // Search for the TOTAL row
  for (i in data) {
    for (j in data[i]) {
      if (String(data[i][j]).search(totalTerm) !== -1) {
        lastNumberRow = Number(i) - 1;
      }
    }
  }

  // Sets the range of cells containing contract data to white
  sheet.getRange(2, 2, lastNumberRow,
                 Number(sheet.getDataRange().getWidth()) - 1)
                 .setBackground("white")
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

  // Sets the colors for each contract type
  var playerOptionColor   = getSetting("player");
  var teamOptionColor     = getSetting("team");
  var vestingOptionColor  = getSetting("vesting");
  var autoContractColor   = getSetting("auto");
  var arbitrationColor    = getSetting("arbitration");
  var minorContractColor  = getSetting("minor");

  var i;

  // Color player option cells
  for (i in cells.playerOptionCells) {
    sheet.getRange(cells.playerOptionCells[i][0],
      cells.playerOptionCells[i][1]).setBackground(playerOptionColor);
  }

  // Color team option cells
  for (i in cells.teamOptionCells) {
    sheet.getRange(cells.teamOptionCells[i][0],
      cells.teamOptionCells[i][1]).setBackground(teamOptionColor);
  }

  // Color vesting option cells
  for (i in cells.vestingOptionCells) {
    sheet.getRange(cells.vestingOptionCells[i][0],
      cells.vestingOptionCells[i][1]).setBackground(vestingOptionColor);
  }

  // Color auto contract cells
  for (i in cells.autoContractCells) {
    sheet.getRange(cells.autoContractCells[i][0],
      cells.autoContractCells[i][1]).setBackground(autoContractColor);
  }

  // Color arbitration contract cells
  for (i in cells.arbitrationCells) {
    sheet.getRange(cells.arbitrationCells[i][0],
      cells.arbitrationCells[i][1]).setBackground(arbitrationColor);
  }

  // Color minor league contract cells
  for (i in cells.minorContractCells) {
    sheet.getRange(cells.minorContractCells[i][0],
      cells.minorContractCells[i][1]).setBackground(minorContractColor);
  }
}

/**
 * Displays the remaining budget by subtracting budget from salary
 */
function remainingBudget() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();

  var budgetTerm = getSetting("budget");
  var remainingTerm = getSetting("remaining");
  var numberFormat = getSetting("format");

  var i, j, lastRow, currentRow;
  var hasBudget = false;
  var response = "";

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
                            "A budget will now be generated.",
                            ui.ButtonSet.OK);

    // Create the budgets
    addBudgets();

    // Select the last row
    currentRow = Number(sheet.getDataRange().getHeight());
  }

  sheet.getRange(currentRow + 1, 1).setValue(remainingTerm);
  sheet.getRange(currentRow + 1, 2, 1,
                 Number(sheet.getDataRange().getWidth()) - 1)
                 .setValue(Utilities.formatString('=MINUS(B%s, B%s)',
                           currentRow, currentRow - 1))
                 .setNumberFormat(numberFormat);
}

/**
 * Prompts the user to enter budgets and returns those values
 */
function getBudgets() {
  var ui = SpreadsheetApp.getUi();
  var budgets = {
      'last': 0,
      'current': 0,
      'next': 0,
      'two': 0
  };

  var result = null;
  var response = "";

  // Set the default button state to OK
  var button = ui.Button.OK;

  // While the button state is set to OK, prompt for responses
  while (button == ui.Button.OK) {
    // Prompt the user for last year's budget
    result = ui.prompt(
      'Let\'s set up your budgets!',
      'What was your previous year\'s budget?',
      ui.ButtonSet.OK_CANCEL
    );
    button = result.getSelectedButton();

    if (button !== ui.Button.OK) {
      break;
    }
    budgets.last = Number(String(result.getResponseText())
                          .replace(/\...$/g, "").replace(/(\D)/g, ""));

    // Prompt the user for this year's budget
    result = ui.prompt(
      'This Year',
      'What is this year\'s projected budget?',
      ui.ButtonSet.OK_CANCEL
    );
    button = result.getSelectedButton();

    if (button !== ui.Button.OK) {
      break;
    }
    budgets.current = Number(String(result.getResponseText())
                             .replace(/\...$/g, "").replace(/(\D)/g, ""));

    // Prompt the user for next year's projected budget
    result = ui.prompt(
      'Next Year',
      'What is your budget projected to be next year?',
      ui.ButtonSet.OK_CANCEL
    );
    button = result.getSelectedButton();

    if (button !== ui.Button.OK) {
      break;
    }
    budgets.next = Number(String(result.getResponseText())
                          .replace(/\...$/g, "").replace(/(\D)/g, ""));

    // Prompt the user for the projected budget in two years
    result = ui.prompt(
      'Two Years Ahead',
      'What is your budget projected to be in two years?',
      ui.ButtonSet.OK_CANCEL
    );
    button = result.getSelectedButton();

    if (button !== ui.Button.OK) {
      break;
    }
    budgets.two = Number(String(result.getResponseText())
                         .replace(/\...$/g, "").replace(/(\D)/g, ""));

    return budgets;
  }

  // If the user cancels, return nothing
  return null;
}

/**
 * Generates a projection of the future team budget
 * by using Google Sheet's TREND formula.
 */
function addBudgets() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var budgetTerm = getSetting("budget");
  var numberFormat = getSetting("format");

  var currentRow = 0;
  var currentCol = 0;
  var budgets = null;
  var i, j;

  var data = sheet.getDataRange().getValues();

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

  if (budgets !== null || budgets !== undefined) {
    // Set the budget values in the cells
    sheet.getRange(currentRow, 2).setValue(budgets.last);
    sheet.getRange(currentRow, 3).setValue(budgets.current);
    sheet.getRange(currentRow, 4).setValue(budgets.next);
    sheet.getRange(currentRow, 5).setValue(budgets.two);

    // Set the TREND formula for the remaining columns
    sheet.getRange(currentRow, 6).setValue(Utilities
          .formatString('=TREND(B%s:E%s, B1:E1, F1:K1)',
          currentRow, currentRow));

    // Set the number formats for the column
    sheet.getRange(currentRow, 2, 1,
                   sheet.getDataRange().getWidth() - 1)
                  .setNumberFormat(numberFormat);
  }
  else {
    ui.alert("Your budget was not set.");
  }
}

/**
 * A function that detects the totals row, and adds
 * SUM formulas for each of the columns.
 */
function addSalaries() {
  var sheet = SpreadsheetApp.getActiveSheet();

  var totalTerm = getSetting("salary");
  var numberFormat = getSetting("format");

  var cell = [];
  var currentRow = 0;
  var currentCol = 0;
  var i, j;

  var data = sheet.getDataRange().getValues();

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
  sheet.getRange(currentRow + 1, currentCol + 2, 1,
                 sheet.getDataRange().getWidth() - 1)
                 .setValue(Utilities.formatString('=SUM(B2:B%s)', currentRow))
                 .setNumberFormat(numberFormat);
}

/**
 * Executes the cleanSalaries function after adding color
 */
function cleanSalariesColor() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  colorCells();
  sheet.activate();
  cleanSalaries();
}

/**
 * A function that cleans and reformats OOTP salary data,
 * thus making it usable within a spreadsheet environment.
 */
function cleanSalaries() {
  var sheet = SpreadsheetApp.getActiveSheet();

  var numberFormat = getSetting("format");
  var freezeRows = getSetting("freeze");

  var result = [];
  var cellText = "";
  var leadingSlice = 0;
  var leadingNumber = "";
  var data = null;
  var i, j;

  // Retrieve only salary numbers
  data = sheet.getRange(2, 2, sheet.getDataRange().getHeight() - 2,
                        sheet.getDataRange().getWidth() - 2)
                        .getValues();

  // If the options sheet states to freeze the rows, freeze them,
  // otherwise, remove the frozen rows
  if (freezeRows == "Yes")
  {
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }
  else {
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
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

        // If the leading number isn't a 0, include the leading number,
        // and then add the appropriate number of zeros
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

  // Apply the results to the default range.
  sheet.getRange(2, 2, sheet.getDataRange().getHeight() - 2,
                 sheet.getDataRange().getWidth() - 2).setValues(result);

  // Set the formatting of the numbers
  sheet.getRange(2, 2, sheet.getDataRange().getHeight() - 2, 10)
                 .setNumberFormat(numberFormat);
}