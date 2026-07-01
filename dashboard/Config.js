/**
 * Config.js
 * Central configuration for the Funding Cockpit dashboard.
 *
 * Secrets (Xero client id/secret, tenant id) are NOT stored here. They live in
 * Script Properties so they never end up in source control. Set them once via the
 * "Cockpit > Setup" menu, or in Project Settings > Script Properties.
 *
 * Everything in CONFIG below is non-secret structural config and is safe to commit.
 */

const CONFIG = {
  // ---- Google Drive: where the budget Gsheets live -------------------------
  // The dashboard walks every project folder under BUDGETS_ROOT_FOLDER_ID and
  // reads the `secured/` and `proposed/` subfolders. Get the id from the Drive
  // URL of the top-level "Budgets" folder: drive.google.com/drive/folders/<ID>
  BUDGETS_ROOT_FOLDER_ID: '10105co6S5qHFSVVg0pb0fkoPidN3ScJZ',
  SECURED_FOLDER_NAME: 'secured',
  PROPOSED_FOLDER_NAME: 'proposed',
  ARCHIVED_FOLDER_NAME: 'archived',
  ARCHIVE_PREFIX: 'Z_ARCH_', // funding sources / projects with this prefix are ignored

  // ---- Budget sheet structure ---------------------------------------------
  BUDGET_TAB: 'Budget',
  TRACKING_TAB: 'Budget, Actual, Forecast Tracking',
  // Columns on the funding-source `Budget` tab (keyed on milestone, not accounts).
  // Two real variants exist and both are supported:
  //   Description, Start, End, Cost, Income, Contribution, Comments, Milestone,
  //     Xero Inventory Item                              (e.g. SPY_26_HAND)
  //   ...same, plus an *Account column                  (e.g. WW_25_TOI)
  // `account` and `project` are optional. When `*Account` is present it's captured
  // for drill-down/reconciliation (a single milestone can span several accounts).
  // `Project` is per line: a line uses its `Project` value when set; if there's no
  // column or the cell is blank, the line belongs to the parent project folder.
  // Matching is case-insensitive and trimmed.
  BUDGET_COLUMNS: {
    description: 'Description',
    start: 'Start',
    end: 'End',
    cost: 'Cost',
    income: 'Income',
    contribution: 'Contribution',
    account: '*Account',         // optional
    milestone: 'Milestone',
    item: 'Xero Inventory Item', // the {SOURCE}_{NNN} product/service code
    project: 'Project'           // optional
  },
  BUDGET_REQUIRED_COLUMNS: ['start', 'end', 'cost'],
  DEFAULT_PROJECT: 'Unallocated',
  GENERAL_PROJECT: 'General',

  // ---- Chart of accounts: ported from create_xero_budget_project.js --------
  REVENUE_ACCOUNTS: ['Grants (102)', 'Project Contract Income (181)'],
  OVERHEAD_ACCOUNT: 'Overhead Allocation (500)',
  DEFERRED_ACCOUNT: 'Unused Donations and Grants with Conditions (835)',
  // Balance-sheet / non-operational accounts excluded from spend + forecast.
  EXCLUDED_ACCOUNTS: [
    'Accounts Payable (800)', 'Accounts Receivable (610)', 'ANZ Term Deposit (605)',
    'Computer Equipment (720)', 'GST (820)', 'Historical Adjustment (840)',
    'Income Tax (830)', 'Inventory (630)',
    'Less Accumulated Depreciation on Computer Equipment (721)',
    'Less Accumulated Depreciation on Office Equipment (711)',
    'less Provision for Doubtful Debts (611)', 'Loan (900)',
    'Office Equipment (710)', 'Owner A Drawings (980)', 'Owner A Funds Introduced (970)',
    'PAYE Payable (825)', 'Payroll Accrual (834)', 'Prepayments (620)',
    'Retained Earnings (960)', 'Suspense (850)', 'Tracking Transfers (877)',
    'Unpaid Expense Claims (801)', 'Visa Prezzy Card (622)',
    'Wages Deductions Payable (816)', 'Wages Payable - Payroll (814)',
    'WILDLIFE.AI TRUST (600)', 'Withholding tax paid (625)'
  ],

  // ---- Xero --------------------------------------------------------------
  XERO: {
    AUTH_URL: 'https://login.xero.com/identity/connect/authorize',
    TOKEN_URL: 'https://identity.xero.com/connect/token',
    API_BASE: 'https://api.xero.com/api.xro/2.0',
    CONNECTIONS_URL: 'https://api.xero.com/connections',
    // Minimal read-only scopes. accounting.transactions.read covers bank
    // transactions + invoices; accounting.settings.read covers the chart of
    // accounts and tracking categories; offline_access enables token refresh.
    SCOPE: 'offline_access accounting.transactions.read accounting.settings.read',
    PROJECT_TRACKING_CATEGORY: 'Projects',
    FUNDING_TRACKING_CATEGORY: 'Funding source'
  },

  // ---- Quarterly forecast layer ------------------------------------------
  // Replaces the per-sheet "Budget, Actual, Forecast Tracking" tab. The GM edits
  // a forward forecast (spend, per milestone per quarter) in the dashboard; it is
  // persisted to a single central "Forecast" Google Sheet.
  // No manual setup needed: if SPREADSHEET_ID is unset (and Script Property
  // COCKPIT_FORECAST_SHEET_ID is empty), the app auto-creates FILE_NAME inside the
  // Budgets root folder on first use and remembers its id. Set SPREADSHEET_ID only
  // to point at a specific existing sheet instead.
  FORECAST: {
    SPREADSHEET_ID: '',          // leave blank to auto-create
    FILE_NAME: 'Cockpit Forecast',
    TAB: 'Forecast',
    // Amount rows have a Quarter plus a Forecast Cost and/or Forecast Income; a
    // milestone Comment is stored on its own row (Quarter blank, Comment filled).
    HEADER: ['Funding Source', 'Milestone', 'Item', 'Quarter', 'Forecast Cost',
             'Forecast Income', 'Comment', 'Updated By', 'Updated At']
  },

  // ---- Financial year -----------------------------------------------------
  // Quarters in the tracking screen follow this financial year. 4 = April start
  // (Apr-Mar), so Q1 = Apr-Jun, Q2 = Jul-Sep, Q3 = Oct-Dec, Q4 = Jan-Mar.
  FINANCIAL_YEAR_START_MONTH: 4,
  // How far ahead the General project view forecasts (quarters past current).
  GENERAL_FORECAST_QUARTERS: 6, // 1.5 years

  // ---- Caching ------------------------------------------------------------
  // The snapshot is stored as a JSON file in Drive (no size ceiling, unlike
  // Script Properties). The file id is kept in a Script Property.
  SNAPSHOT_FILE_NAME: 'cockpit_snapshot.json',
  SNAPSHOT_FILE_ID_PROPERTY: 'COCKPIT_SNAPSHOT_FILE_ID',
  REFRESH_TRIGGER_HOURS: 6
};

/** Forecast spreadsheet id from Config or Script Property. */
function getForecastSheetId() {
  return getSecret('COCKPIT_FORECAST_SHEET_ID') || CONFIG.FORECAST.SPREADSHEET_ID;
}

/** Read a secret from Script Properties (returns '' if unset). */
function getSecret(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

/** Store a secret in Script Properties. */
function setSecret(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}
