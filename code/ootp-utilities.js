/**
 * Adds "OOTP" to the menu, and provides access to various functions
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu('OOTP')
    .addItem('Simple Format', 'simpleFormat')
    .addItem('Simple Format \(Add Color\)', 'simpleFormatColor')
    .addSeparator()
    .addItem('Expert Format', 'expertFormat')
    .addItem('Expert Format \(Add Color\)', 'expertFormatColor')
    .addSeparator()
    .addItem('(Re)Compute Budgets', 'addBudgets')
    .addSeparator()
    .addItem('Add/Update Color \(Do Before Format\)', 'colorCells')
    .addItem('Remove Color', 'removeColor')
    .addSeparator()
    .addItem('Add/Reset Settings Sheet', 'generateSettingsSheet')
  .addToUi()
}

/**
 * Formats the data with helpful prompts and a simpler expense model
 */
function simpleFormat() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Locate the settings sheet if it exists
  var settingsSheet = spreadsheet.getSheetByName("settings");

  // If settings sheet doesn't exist, create it
  if (settingsSheet == null) {
    generateSettingsSheet();
  }

  // Reactivate the original sheet
  sheet.activate();

  var formatted = checkFormatting(sheet);

  // If the data was not previously formatted, run normally
  if (formatted == false) {
    // Clean the salaries
    cleanSalaries();

    // Render remaining
    remainingBudget(true);

    // Display Payroll Total
    addSalaries();

    // Display Other Expenses
    addOtherExpenses(sheet, true);

    // Display Other Income
    addOtherIncome(sheet);

    // Display Budget
    addBudgets();
  }
  // If the data was previously formatted, only clean the salary data
  else {
    cleanSalaries();
  }
}

/**
 * Adds color before formatting the data with the simplified model
 */
function simpleFormatColor() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Locate the settings sheet if it exists
  var settingsSheet = spreadsheet.getSheetByName("settings");

  // If settings sheet doesn't exist, create it
  if (settingsSheet == null) {
    generateSettingsSheet();
  }

  // Reactivate the original sheet
  sheet.activate();

  colorCells();
  simpleFormat();
}

/**
 * Formats the data with expert options and an expansive expense model
 */
function expertFormat() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Locate the settings sheet if it exists
  var settingsSheet = spreadsheet.getSheetByName("settings");

  // If settings sheet doesn't exist, create it
  if (settingsSheet == null) {
    generateSettingsSheet();
  }

  // Reactivate the original sheet
  sheet.activate();

  var formatted = checkFormatting(sheet);

  // If the data was not previously formatted, run normally
  if (formatted == false) {
    // Clean the salaries
    cleanSalaries();

    // Render remaining
    remainingBudget();

    // Display Payroll Total
    addSalaries();

    // Display Other Expenses
    addOtherExpenses(sheet, false);

    // Display Other Income
    addOtherIncome(sheet);

    // Display Budget
    addBudgets();
  }
  // If the data was previously formatted, only clean the salary data
  else {
    cleanSalaries();
  }
}

/**
 * Adds color before formatting the data with the expert model
 */
function expertFormatColor() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Locate the settings sheet if it exists
  var settingsSheet = spreadsheet.getSheetByName("settings");

  // If settings sheet doesn't exist, create it
  if (settingsSheet == null) {
    generateSettingsSheet();
  }

  // Reactivate the original sheet
  sheet.activate();

  colorCells();
  expertFormat();
}

/**
 * Checks if a format of the data has been previously initiated
 */
function checkFormatting(sheet) {
  var data = sheet.getDataRange().getValues();
  var totalRows = sheet.getDataRange().getHeight();

  var remainingTerm = getSetting("remainingID");
  var i = null;

  // Search for the remaining term, and return it if it's found
  for (i in data) {
    if (data[i][0] == remainingTerm) {
      return Number(i);
    }
  }

  return false;
}

/**
 * Retrieves the first row of summary items
 */
function getFirstSummaryRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var totalRows = sheet.getDataRange().getHeight();

  var totalDefault = "TOTAL";
  var totalTerm = getSetting("salaryID");
  var remainingTerm = getSetting("remainingID");

  var firstCellRow = null;

  var i;

  // Search through the first column to find the firstCell
  for (i in data) {
    if (
        data[i][0] == totalDefault ||
        data[i][0] == totalTerm ||
        data[i][0] == remainingTerm
      ) {
        return firstCellRow = Number(i);
      }
  }

  // If the firstCell is not found, assume the last cell is it
  if (firstCellRow == null) {
    firstCellRow = Number(totalRows) + 1;
  }

  return firstCellRow;
};

/**
 * Adds other income summary field
 */
function addOtherIncome(sheet) {
  var lastRow = sheet.getDataRange().getHeight() + 1;
  var lastCol = sheet.getDataRange().getWidth();

  var numberFormat = getSetting("format");
  var incomeID = getSetting("incomeID");
  var incomeColor = getSetting("incomeColor");

  sheet.getRange(lastRow, 1, 1, 1).setValue(incomeID);
  sheet.getRange(lastRow, 2, 1, lastCol - 1)
                 .setNumberFormat(numberFormat);

  // Set the color
  sheet.getRange(lastRow, 1, 1, lastCol).setBackground(incomeColor);
}

/**
 * Adds other expense fields for expert formatting
 */
function addOtherExpenses(sheet, simplified) {
  var lastRow = getFirstSummaryRow() + 2;
  var lastCol = sheet.getDataRange().getWidth();

  var numberFormat = getSetting("format");
  var expensesID = getSetting("expensesID");
  var staffID = getSetting("staffID");
  var scoutingID = getSetting("scoutingID");
  var draftID = getSetting("draftID");
  var playerDevID = getSetting("playerDevID");
  var miscID = getSetting("miscID");
  var expensesColor = getSetting("expensesColor");
  var expensesIndividualColor = getSetting("expensesIndividualColor");

  if (simplified == true) {
    sheet.getRange(lastRow + 1, 1, 1, 1).setValue(expensesID);
    // Set the color and format
    sheet.getRange(lastRow + 1, 1, 1, lastCol).setBackground(expensesColor);
    sheet.getRange(lastRow + 1, 2, 1, lastCol - 1)
                   .setNumberFormat(numberFormat);
  }
  else {
    // Insert other expenses
    sheet.getRange(lastRow + 1, 1, 1, 1).setValue(staffID);
    sheet.getRange(lastRow + 2, 1, 1, 1).setValue(scoutingID);
    sheet.getRange(lastRow + 3, 1, 1, 1).setValue(draftID);
    sheet.getRange(lastRow + 4, 1, 1, 1).setValue(playerDevID);
    sheet.getRange(lastRow + 5, 1, 1, 1).setValue(miscID);

    // Set the color and format
    sheet.getRange(lastRow + 1, 1, 5, lastCol)
                   .setBackground(expensesIndividualColor);
    sheet.getRange(lastRow + 1, 2, 5, lastCol - 1)
                   .setNumberFormat(numberFormat);

    sheet.getRange(lastRow + 6, 1, 1, 1).setValue(expensesID);
    sheet.getRange(lastRow + 6, 2, 1, lastCol - 1)
                   .setValue(Utilities.formatString('=SUM(B%s:B%s)',
                                                    lastRow + 1, lastRow + 5))
                   .setNumberFormat(numberFormat);

    // Set the color
    sheet.getRange(lastRow + 6, 1, 1, lastCol).setBackground(expensesColor);
  }
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
    case "format":
      cell = sheet.getRange(5, 2, 1, 1);
      break;
    case "player":
      cell = sheet.getRange(8, 2, 1, 1);
      break;
    case "team":
      cell = sheet.getRange(9, 2, 1, 1);
      break;
    case "vesting":
      cell = sheet.getRange(10, 2, 1, 1);
      break;
    case "auto":
      cell = sheet.getRange(11, 2, 1, 1);
      break;
    case "arbitration":
      cell = sheet.getRange(12, 2, 1, 1);
      break;
    case "minor":
      cell = sheet.getRange(13, 2, 1, 1);
      break;
    case "remainingColor":
      cell = sheet.getRange(16, 2, 1, 1);
      break;
    case "salaryColor":
      cell = sheet.getRange(17, 2, 1, 1);
      break;
    case "expensesIndividualColor":
      cell = sheet.getRange(18, 2, 1, 1);
      break;
    case "expensesColor":
      cell = sheet.getRange(19, 2, 1, 1);
      break;
    case "incomeColor":
      cell = sheet.getRange(20, 2, 1, 1);
      break;
    case "budgetColor":
      cell = sheet.getRange(21, 2, 1, 1);
      break;
    case "remainingID":
      cell = sheet.getRange(24, 2, 1, 1);
      break;
    case "salaryID":
      cell = sheet.getRange(25, 2, 1, 1);
      break;
    case "staffID":
      cell = sheet.getRange(26, 2, 1, 1);
      break;
    case "scoutingID":
      cell = sheet.getRange(27, 2, 1, 1);
      break;
    case "draftID":
      cell = sheet.getRange(28, 2, 1, 1);
      break;
    case "playerDevID":
      cell = sheet.getRange(29, 2, 1, 1);
      break;
    case "miscID":
      cell = sheet.getRange(30, 2, 1, 1);
      break;
    case "expensesID":
      cell = sheet.getRange(31, 2, 1, 1);
      break;
    case "incomeID":
      cell = sheet.getRange(32, 2, 1, 1);
      break;
    case "budgetID":
      cell = sheet.getRange(33, 2, 1, 1);
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

  //

  cell = sheet.getRange(4, 1, 1, 1).setValue("Number Format")
         .setFontWeight("Bold");
  cell = sheet.getRange(4, 2, 1, 1).setValue("For Help, Visit:")
         .setFontWeight("Bold");
  cell = sheet.getRange(4, 3, 1, 1)
         .setValue("https://support.google.com/docs/answer/56470?hl=en");

  cell = sheet.getRange(5, 1, 1, 1).setValue("Format");
  cell = sheet.getRange(5, 2, 1, 1).setValue("$#,##0_)");

  //

  cell = sheet.getRange(7, 1, 1, 1).setValue("Contract Color Options")
         .setFontWeight("Bold");
  cell = sheet.getRange(7, 2, 1, 1).setValue("Color Value \(#XXXXXX\)")
         .setFontWeight("Bold");

  cell = sheet.getRange(8, 1, 1, 1).setValue("Player Option");
  cell = sheet.getRange(8, 2, 1, 1).setValue("#FB8072");

  cell = sheet.getRange(9, 1, 1, 1).setValue("Team Option");
  cell = sheet.getRange(9, 2, 1, 1).setValue("#B7D2FF");

  cell = sheet.getRange(10, 1, 1, 1).setValue("Vesting Option");
  cell = sheet.getRange(10, 2, 1, 1).setValue("#BEBADA");

  cell = sheet.getRange(11, 1, 1, 1).setValue("Auto Contract");
  cell = sheet.getRange(11, 2, 1, 1).setValue("#8DD3C7");

  cell = sheet.getRange(12, 1, 1, 1).setValue("Arbitration");
  cell = sheet.getRange(12, 2, 1, 1).setValue("#FFFFB3");

  cell = sheet.getRange(13, 1, 1, 1).setValue("Minor League");
  cell = sheet.getRange(13, 2, 1, 1).setValue("#ECECEC");

  //

  cell = sheet.getRange(15, 1, 1, 1).setValue("Summary Color Options")
         .setFontWeight("Bold");
  cell = sheet.getRange(15, 2, 1, 1).setValue("Color Value \(#XXXXXX\)")
         .setFontWeight("Bold");

  cell = sheet.getRange(16, 1, 1, 1).setValue("Remaining Total");
  cell = sheet.getRange(16, 2, 1, 1).setValue("#DAEBD4");

  cell = sheet.getRange(17, 1, 1, 1).setValue("Salary Total");
  cell = sheet.getRange(17, 2, 1, 1).setValue("#FA8176");

  cell = sheet.getRange(18, 1, 1, 1).setValue("Other Individual Expenses");
  cell = sheet.getRange(18, 2, 1, 1).setValue("#EBD2DD");

  cell = sheet.getRange(19, 1, 1, 1).setValue("Other Expenses Total");
  cell = sheet.getRange(19, 2, 1, 1).setValue("#EAB8B8");

  cell = sheet.getRange(20, 1, 1, 1).setValue("Other Income Total");
  cell = sheet.getRange(20, 2, 1, 1).setValue("#76CF77");

  cell = sheet.getRange(21, 1, 1, 1).setValue("Budget Total");
  cell = sheet.getRange(21, 2, 1, 1).setValue("#CADBF8");


  //

  cell = sheet.getRange(23, 1, 1, 1).setValue("Row Text Descriptions")
         .setFontWeight("Bold");
  cell = sheet.getRange(23, 2, 1, 1).setValue("Text")
         .setFontWeight("Bold");

  cell = sheet.getRange(24, 1, 1, 1).setValue("Remaining Total");
  cell = sheet.getRange(24, 2, 1, 1).setValue("REMAINING PAYROLL");

  cell = sheet.getRange(25, 1, 1, 1).setValue("Salary Total");
  cell = sheet.getRange(25, 2, 1, 1).setValue("PAYROLL TOTAL");

  cell = sheet.getRange(26, 1, 1, 1).setValue("Staff Expenses");
  cell = sheet.getRange(26, 2, 1, 1).setValue("STAFF EXPENSES");

  cell = sheet.getRange(27, 1, 1, 1).setValue("Scouting Expenses");
  cell = sheet.getRange(27, 2, 1, 1).setValue("SCOUTING EXPENSES");

  cell = sheet.getRange(28, 1, 1, 1).setValue("Draft Expenses");
  cell = sheet.getRange(28, 2, 1, 1).setValue("DRAFT EXPENSES");

  cell = sheet.getRange(29, 1, 1, 1).setValue("Player Development Expenses");
  cell = sheet.getRange(29, 2, 1, 1).setValue("PLAYER DEV EXPENSES");

  cell = sheet.getRange(30, 1, 1, 1).setValue("Miscellaneous Player Expenses");
  cell = sheet.getRange(30, 2, 1, 1).setValue("MISC PLAYER EXPENSES");

  cell = sheet.getRange(31, 1, 1, 1).setValue("Other Expenses Total");
  cell = sheet.getRange(31, 2, 1, 1).setValue("OTHER EXPENSES");

  cell = sheet.getRange(32, 1, 1, 1).setValue("Other Income Total");
  cell = sheet.getRange(32, 2, 1, 1).setValue("OTHER INCOME");

  cell = sheet.getRange(33, 1, 1, 1).setValue("Budget Total");
  cell = sheet.getRange(33, 2, 1, 1).setValue("BUDGET");

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

  var lastNumberRow = getFirstSummaryRow() - 1;

  // Sets the range of cells containing contract data to white
  sheet.getRange(2, 2, lastNumberRow,
                 Number(sheet.getDataRange().getWidth()) - 1)
                 .setBackground("white");
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
      if (String(data[i][j]).search(/MiLC/g) !== -1) {
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
function remainingBudget(simplified) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();

  var budgetTerm = getSetting("budgetID");
  var remainingTerm = getSetting("remainingID");
  var numberFormat = getSetting("format");
  var remainingColor = getSetting("remainingColor");

  var i, j, lastRow, currentRow;
  var hasBudget = false;
  var response = "";

  lastRow = getFirstSummaryRow();

  if (simplified == true) {
    if (sheet.getRange(lastRow, 1).getValue() == "") {
      sheet.getRange(lastRow, 1).setValue(remainingTerm)
                   .setBackground(remainingColor);

      sheet.getRange(lastRow, 2, 1,
                   Number(sheet.getDataRange().getWidth()) - 1)
                   .setValue(Utilities.formatString('=SUM(B%s - B%s - B%s + B%s)',
                             lastRow + 5, lastRow + 2, lastRow + 3, lastRow + 4))
                   .setBackground(remainingColor)
                   .setNumberFormat(numberFormat);
    }
    else {
      sheet.insertRowBefore(lastRow + 1);

      sheet.getRange(lastRow + 1, 1).setValue(remainingTerm)
                   .setBackground(remainingColor);

      sheet.getRange(lastRow + 1, 2, 1,
                   Number(sheet.getDataRange().getWidth()) - 1)
                   .setValue(Utilities.formatString('=SUM(B%s - B%s - B%s + B%s)',
                             lastRow + 5, lastRow + 2, lastRow + 3, lastRow + 4))
                   .setBackground(remainingColor)
                   .setNumberFormat(numberFormat);
    }
  }
  else {
    if (sheet.getRange(lastRow, 1).getValue() == "") {
      sheet.getRange(lastRow, 1).setValue(remainingTerm)
                   .setBackground(remainingColor);

      sheet.getRange(lastRow, 2, 1,
                   Number(sheet.getDataRange().getWidth()) - 1)
                   .setValue(Utilities.formatString('=SUM(B%s - B%s - B%s + B%s)',
                             lastRow + 10, lastRow + 2, lastRow + 8, lastRow + 9))
                   .setBackground(remainingColor)
                   .setNumberFormat(numberFormat);
    }
    else {
      sheet.insertRowBefore(lastRow + 1);

      sheet.getRange(lastRow + 1, 1).setValue(remainingTerm)
                   .setBackground(remainingColor);

      sheet.getRange(lastRow + 1, 2, 1,
                   Number(sheet.getDataRange().getWidth()) - 1)
                   .setValue(Utilities.formatString('=SUM(B%s - B%s - B%s + B%s)',
                             lastRow + 10, lastRow + 2, lastRow + 8, lastRow + 9))
                   .setBackground(remainingColor)
                   .setNumberFormat(numberFormat);
    }
  }
}

/**
 * Prompts the user to enter budgets and returns those values
 */
function getBudgets() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var budgets = {
      'current': 0,
      'next': 0,
      'two': 0
  };

  // Grab the correct years to be used for the prompts
  var thisYear  = String(sheet.getRange(1, 2).getValue());
  var nextYear  = String(sheet.getRange(1, 3).getValue());
  var twoYear   = String(sheet.getRange(1, 4).getValue());

  if (
    thisYear == "" ||
    nextYear == "" ||
    twoYear == ""
  ) {
    ui.alert("Your sheet has not been formatted correctly!");
    return null;
  }

  var result = null;
  var response = "";

  // Set the default button state to OK
  var button = ui.Button.OK;

  // While the button state is set to OK, prompt for responses
  while (button == ui.Button.OK) {
    // Prompt the user for this year's budget
    result = ui.prompt(
      thisYear + ' Budget',
      'What is the budget for ' + thisYear + '?',
      ui.ButtonSet.OK_CANCEL
    );
    button = result.getSelectedButton();

    if (button !== ui.Button.OK) {
      return null;
    }
    budgets.current = Number(String(result.getResponseText())
                             .replace(/\...$/g, "").replace(/(\D)/g, ""));

    // Prompt the user for next year's projected budget
    result = ui.prompt(
      nextYear + ' Budget',
      'What is the budget for ' + nextYear + '?',
      ui.ButtonSet.OK_CANCEL
    );
    button = result.getSelectedButton();

    if (button !== ui.Button.OK) {
      return null;
    }
    budgets.next = Number(String(result.getResponseText())
                          .replace(/\...$/g, "").replace(/(\D)/g, ""));

    // Prompt the user for the projected budget in two years
    result = ui.prompt(
      twoYear + ' Budget',
      'What is the budget for ' + twoYear + '?',
      ui.ButtonSet.OK_CANCEL
    );
    button = result.getSelectedButton();

    if (button !== ui.Button.OK) {
      return null;
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

  var budgetTerm = getSetting("budgetID");
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
    sheet.getRange(currentRow, 1).setValue(budgetTerm).setBackground("#cadbf8");
    sheet.getRange(currentRow, 2, 1, sheet.getDataRange().getWidth() - 1)
                   .setBackground("#cadbf8")
                   .setNumberFormat(numberFormat);
  }

  // Clear the budget row to prevent problems with formulas
  sheet.getRange(currentRow, 2, 1, sheet.getDataRange().getWidth() - 1)
                 .clearContent();

  // Retrieve the budgets from the user
  budgets = getBudgets();

  if (budgets == null) {
    return null;
  }
  else if (budgets !== null || budgets !== undefined) {
    // Set the budget values in the cells
    sheet.getRange(currentRow, 2).setValue(budgets.current);
    sheet.getRange(currentRow, 3).setValue(budgets.next);
    sheet.getRange(currentRow, 4).setValue(budgets.two);

    // Set the TREND formula for the remaining columns
    sheet.getRange(currentRow, 5).setValue(Utilities
          .formatString('=TREND(B%s:D%s, B1:D1, E1:K1)',
          currentRow, currentRow));

    // Set the number formats for the column
    sheet.getRange(currentRow, 2, 1,
                   sheet.getDataRange().getWidth() - 1)
                  .setBackground("#cadbf8")
                  .setNumberFormat(numberFormat);
  }
  else {
    return null;
  }
}

/**
 * A function that detects the totals row, and adds
 * SUM formulas for each of the columns.
 */
function addSalaries() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();

  var totalTerm = getSetting("salaryID");
  var numberFormat = getSetting("format");
  var salaryColor = getSetting("salaryColor");

  var cell = [];
  var currentRow = 0;
  var currentCol = 0;
  var i, j;

  currentRow = getFirstSummaryRow() + 2;

  // If the salary term does not match the totalTerm in the options
  // set it to the totalTerm or create it
  if (String(sheet.getRange(currentRow, 1).getValue()) == "TOTAL") {
    // ui.alert("The salary total is being modified to your custom settings.");
    sheet.getRange(currentRow, 1).setValue(totalTerm);
  }
  else if (
    sheet.getRange(currentRow, 1).getValue() == "" ||
    sheet.getRange(currentRow, 1).getValue() == 0
  ) {
    //ui.alert("The salary total was not found and is \
    //          being created based on your custom settings");
    sheet.getRange(currentRow, 1).setValue(totalTerm);
  }
  else {
    ui.alert("Critical error. Check out the OOTP Utilities Visual Guide!");
    return null;
  }

  sheet.getRange(currentRow, 1).setBackground(salaryColor);

  // Inserts the SUM formulas in the row that represents totals
  sheet.getRange(currentRow, currentCol + 2, 1,
                 sheet.getDataRange().getWidth() - 1)
                 .setValue(Utilities.formatString('=SUM(B2:B%s)', currentRow - 2))
                 .setBackground(salaryColor)
                 .setNumberFormat(numberFormat);
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
  data = sheet.getRange(2, 2, getFirstSummaryRow() - 1,
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
  sheet.getRange(2, 2, getFirstSummaryRow() - 1,
                 sheet.getDataRange().getWidth() - 2).setValues(result);

  // Set the formatting of the numbers
  sheet.getRange(2, 2, getFirstSummaryRow() - 1, 10)
                 .setNumberFormat(numberFormat);
}
