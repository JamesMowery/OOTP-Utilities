/**
 * Adds "OOTP" to the menu, and provides access to various functions
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Salary Cleaner', functionName: 'cleanSalaries'},
    {name: 'Salary Totals', functionName: 'addSalaries'}
  ];
  spreadsheet.addMenu('OOTP', menuItems);
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
  for (i in data) {
    for (j in data[i]) {
      if (String(data[i][j]).search(totalTerm) !== -1) {

        // When found, store the row and column values
        var currentRow = Number(i);
        var currentCol = Number(j);
      }
    }
  }

  // Inserts the SUM formulas in the row that represents totals
  sheet.getRange(currentRow + 1, currentCol + 2,
                 1, sheet.getDataRange().getWidth() - 1)
                 .setValue(Utilities.formatString('=SUM(B2:B%s)', currentRow));

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
    // sheet.getActiveRange().setValues(result).setNumberFormat("$#,##0_)");
  }
  else {
    sheet.getRange(2, 2, sheet.getDataRange().getHeight() - 2,
                   sheet.getDataRange().getWidth() - 2).setValues(result);
  }

  // Set the formatting of the numbers
  sheet.getRange(2, 2, sheet.getDataRange().getHeight() - 2,
                 10).setNumberFormat("$#,##0_)");

}
