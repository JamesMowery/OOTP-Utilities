# OOTP Utilities by James Mowery

## About This Project

This GitHub repository contains OOTP Utilities, which is a collection of code snippets (or otherwise) designed to help players of the game Out of the Park (OOTP) baseball.

## Google Sheets Utilities

Included are several Google Sheets custom functions that strips and correctly formats data from the Salaries page extracted from Out of the Park Baseball.

**Instructions:**

Navigate to Team > Front Office > Salaries.


Use the "Open In External Browser" option.


Copy from the top left headers (Name/Years) to the end of the "TOTAL" section.

Paste this information into a new Google Sheet.

Within this Google Sheet, navigate to Tools > Script Editor

Copy and paste the code from within ootp-utilities.js.

Refresh (or re-open) your Google Sheet.

You will now see a new menu called "OOTP".

### Salary Cleaner

Salary Cleaner strips out extra data and properly formats the resulting text into data that can be easily manipulated and utilized within a spreadsheet.

* Removed parenthetical information
* Removes punctuation
* Expands numeric shorthands (m and k) to proper amounts
* Appropriately formats the results for display and modification
* Enables the usage of formulas on the results

### Salary Totals

Salary Totals searches the spreadsheet for the word "TOTAL", and then populates that row with a SUM formula that generates dynamic totals (as opposed to the static totals that OOTP generates).

### Credits

Find more information about the author, James Mowery, at his homepage: http://mowery.co/

Special thanks to Out of the Park Developments for creating a wonderful baseball simulation that I've spent countless hours enjoying.
