# Project Reports & Quarterly Budgets

This folder contains scripts related to generating project reports and aggregating budgets for Xero.

## `create_quarterly_budgets.js`

This script parses funding source budget sheets for the four main projects and generates a summary sheet `YYQX_PROJECT` that contains the aggregated budget amounts formatted monthly. It also simulates sending these budgets to Xero via the API.

### Configuration

- **Budgets Folder ID**: `10105co6S5qHFSVVg0pb0fkoPidN3ScJZ`
- **Projects**: Spyfish Aotearoa (SPY), General (GEN), Wild About AI (WAA), Wildlife Watcher (WLW).
- **Target Sheet**: The script looks specifically for the tab named `Budget, Actual, Forecast Tracking`.
- **Extracted Columns**: G, I, J, K, L (indices 6, 8, 9, 10, 11).

### Logic Breakdown

1. **Date Parsing**: The script uses an April 1st Financial Year start date. It converts the date found in row 4 (e.g. `30-Jun-2026`) into a FY Quarter (e.g., `Q1`) and maps the values into the middle month of that quarter (e.g., `May 2026`).
2. **Account Validation**: Pulls the master list of Xero Accounts from the main chart of accounts spreadsheet (`1VHtZsZRzJJ29tt3SRDXnIP6ebKChoDtlHqyV5Ua-3Yg`).
3. **Summary Sheet**: For every project and quarter, it creates/updates a Google Spreadsheet named `YYQX_PROJECT` (e.g. `26Q2_SPY`).
4. **Funding Source Tab**: Within the summary sheet, a tab is generated for each funding source (e.g. `SPY_26_HAND`), containing column A prepopulated with all valid Xero Accounts, and the middle-month column (e.g. `Aug 2026`) with the aggregated totals.

### Xero Budgets API

Currently, the native Xero API **does not support POST/PUT requests for Budgets**. Therefore, `pushBudgetToXero` acts as a dry-run placeholder. It logs the exact JSON payload that *would* be pushed if the endpoint were available, allowing for easy manual auditing or integration with a 3rd party tool if needed.

### Usage

To run this in Google Apps Script:
1. Open Google Apps Script.
2. Add `create_quarterly_budgets.js` to your project.
3. Make sure the Google services required are enabled (DriveApp, SpreadsheetApp).
4. Run the `generateAndUploadQuarterlyBudgets` function.
