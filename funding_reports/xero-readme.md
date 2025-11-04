# Xero to Google Sheets Integration

Automatically sync Xero transaction data to Google Sheets with tracking category filtering.

## Features

- 🔐 OAuth 2.0 authentication with Xero
- 📊 Fetches journal transactions from Xero API
- 🎯 Filters by tracking category (uses Google Sheet filename)
- 🚫 Excludes specific account codes (835, 600, 800, 820)
- 📅 Customizable date range (from specific start date to today)
- 🔄 One-click sync via custom Google Sheets menu
- 💾 Updates "Xero Transactions" sheet automatically

## What This Script Does

This integration pulls transaction data from Xero and populates a Google Sheet with:
- Transaction dates
- Journal numbers and references
- Account codes and names
- Debit/Credit amounts
- Tracking category information
- Tax and gross amounts

The script filters transactions based on:
1. **Tracking Category**: "Funding source" with value matching your Google Sheet filename
2. **Date Range**: From the date in cell B3 of "Budget, Actual, Forecast Tracking" sheet to today
3. **Account Exclusions**: Excludes accounts 835, 600, 800, and 820

## Prerequisites

- A Google account with access to Google Sheets
- A Xero account with API access
- Basic familiarity with Google Apps Script

## Installation

### 1. Set Up Xero OAuth App

1. Go to [Xero Developer Portal](https://developer.xero.com/app/manage)
2. Click **"New app"**
3. Fill in the required details:
   - **App name**: Choose any name (e.g., "Google Sheets Integration")
   - **Company/application URL**: Your website or placeholder URL
   - **Redirect URI**: Leave blank for now (you'll add this in step 4)
4. Click **"Create app"**
5. Copy and save your **Client ID** and **Client Secret** (you'll need these later)

### 2. Create Your Google Sheet

1. Create a new Google Sheet or open an existing one
2. Name the file to match your Xero tracking category value (e.g., "WW_25_TOI")
3. Create a sheet named **"Budget, Actual, Forecast Tracking"**
4. In cell **B3** of that sheet, enter your start date (e.g., `2024-01-01`)

### 3. Set Up Apps Script

1. In your Google Sheet, go to **Extensions** > **Apps Script**
2. Delete any existing code in the editor
3. Copy the code from **[loader-script.js](loader-script.js)** in this repository
4. Paste it into the Apps Script editor
5. Save the project (Ctrl+S or Cmd+S)
6. Give it a name (e.g., "Xero Integration")

### 4. Install OAuth2 Library

1. In the Apps Script editor, click the **"+"** icon next to **Libraries** (left sidebar)
2. Enter this Script ID: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`
3. Click **"Look up"**
4. Select the latest version
5. Click **"Add"**

### 5. Deploy as Web App

1. In Apps Script, click **"Deploy"** > **"New deployment"**
2. Click the gear icon next to "Select type"
3. Choose **"Web app"**
4. Configure settings:
   - **Description**: "Xero OAuth Handler"
   - **Execute as**: "Me (your@email.com)"
   - **Who has access**: "Anyone"
5. Click **"Deploy"**
6. Click **"Authorize access"** and grant permissions
7. Click **"Done"**

### 6. Configure Credentials

1. In Apps Script, click the **gear icon** (Project Settings) on the left sidebar
2. Copy your **Script ID**
3. Create your redirect URI in this format:
   ```
   https://script.google.com/macros/d/YOUR_SCRIPT_ID_HERE/usercallback
   ```
4. Go back to your Xero app settings and add this redirect URI
5. Save the Xero app settings

6. In your Apps Script loader script, update the `PRIVATE_CONFIG` section:
   ```javascript
   const PRIVATE_CONFIG = {
     CLIENT_ID: 'your_xero_client_id',
     CLIENT_SECRET: 'your_xero_client_secret',
     REDIRECT_URI: 'https://script.google.com/macros/d/YOUR_SCRIPT_ID/usercallback',
     GITHUB_SCRIPT_URL: 'https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/xero-integration.js'
   };
   ```
7. Save the script

## Usage

### First Time Setup

1. **Reload your Google Sheet**
2. You should see a new **"Xero Sync"** menu in the menu bar (next to Help)
3. Click **"Xero Sync"** > **"1. Authorize Xero"**
4. Click the authorization link in the popup
5. Log in to Xero and authorize the connection
6. Close the authorization window

### Syncing Transactions

1. Click **"Xero Sync"** > **"2. Update Transactions"**
2. Wait for the sync to complete (you'll see a progress message)
3. A success message will show the number of transactions synced
4. Check the **"Xero Transactions"** sheet for your data

### Clearing Authorization

If you need to re-authorize or disconnect:
- Click **"Xero Sync"** > **"Clear Authorization"**

## Configuration Options

You can customize the behavior by modifying the `CONFIG` object in `xero-integration.js`:

```javascript
let CONFIG = {
  SHEET_NAME: 'Xero Transactions',              // Name of output sheet
  TRACKING_CATEGORY_NAME: 'Funding source',     // Tracking category to filter
  DATE_SHEET_NAME: 'Budget, Actual, Forecast Tracking', // Sheet with start date
  DATE_CELL: 'B3',                              // Cell containing start date
  EXCLUDED_ACCOUNT_CODES: ['835', '600', '800', '820'] // Accounts to exclude
};
```

## Troubleshooting

### Menu Doesn't Appear
- Go to Apps Script and run the `onOpen` function manually
- Reload your Google Sheet

### Authorization Fails
- Verify your Client ID and Client Secret are correct
- Check that the Redirect URI in Xero matches exactly
- Make sure you deployed the script as a web app

### No Transactions Appear
- Check that your Google Sheet filename matches your Xero tracking category value exactly
- Verify cell B3 in "Budget, Actual, Forecast Tracking" contains a valid date
- Check Apps Script logs (View > Logs) for error messages
- Ensure the tracking category "Funding source" exists in Xero with the correct value

### Permission Errors
- Run any function once from Apps Script to trigger authorization
- Make sure you granted all requested permissions

## File Structure

```
.
├── README.md                  # This file
├── xero-integration.js        # Main integration code (public on GitHub)
└── loader-script.js           # Loader script with credentials (private - not committed)
```

## Security Notes

- **Never commit your `loader-script.js` file** or any file containing your Client ID, Client Secret, or Redirect URI
- Add `loader-script.js` to your `.gitignore`
- Only the `xero-integration.js` file should be public
- Your credentials are stored only in your Google Apps Script project

## License

MIT License - Feel free to use and modify as needed.

## Support

For issues or questions:
1. Check the [Xero API Documentation](https://developer.xero.com/documentation/)
2. Review Apps Script logs for error messages
3. Open an issue in this repository

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

---

**Note**: This integration uses Xero's Journals API endpoint, which provides access to all general ledger transactions. Custom reports are not directly available through the Xero API.