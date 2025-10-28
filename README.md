# Monthly-Budget-Tracker (Apple Shortcut + Excel VBA)
A smart Excel budget tracker powered by Apple Shortcuts and VBA automation.

**What’s inside:**
- Add Transaction.shortcut (Shortcut for Mac to capture entries (Amount, Category, Remarks, From, To)
- Addexpensetoexcel (AppleScript, no extra setup needed)
- Monthly Budget Tracker.xlsm (Template workbook structure)
- UpdateBudget.bas (Contains `UpdateBudgetAll`, which calculates totals, formats, used %, remaining balance)
- UpdateAssets.bas (Contains `UpdateAssetsFromSelectedMonth`, which updates your Assets sheet based on logs)
- Worksheet Change.cls (Worksheet event (`Worksheet_Change`), which triggers `UpdateBudgetAll` automatically after each new entry

**Quick Start:**
1. Download all files in this repository.
2. Copy `Monthly Budget Tracker.xlsm` to your folder.
3. In Excel:
   - Open the workbook.
   - Open the VBA editor.
   - Import:
     - `UpdateBudget.bas`
     - `UpdateAssets.bas`
     - `Worksheet Change.cls` worksheet module for the Template and 2025-Oct                                   sheet)
   - Save and close the VBA editor.
4. In Shortcuts:
   - Import `Add Transaction.shortcut` (double-click on Mac)
   - Ensure your workbook name matches `Monthly Budget Tracker.xlsm` and amend        the path and wortsheet name `2025-Oct`)
6. Run the Shortcut:
   - The script will add a new row to the monthly sheet (e.g. `2025-Oct`),  
     then the worksheet event automatically runs `UpdateBudgetAll` to refresh totals.

**Assets Updater:**
Use `UpdateAssetsFromSelectedMonth` to refresh balances:
1. In the **Assets** sheet, enter:
   - `B1`: Year (e.g. `2025`)
   - `B2`: Month abbreviation (e.g. `Oct`)
   - Column C: Manually input your opening balances at the beginning of               the month.
     (These will stay fixed and serve as the baseline for calculating changes.)
2. Run the macro with the Updated and Saved button.
3. Balances update automatically from the log sheet.
