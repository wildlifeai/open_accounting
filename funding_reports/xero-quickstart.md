# Xero to Google Sheets Integration

Automatically sync Xero transaction data into Google Sheets with a scalable, multi-sheet architecture.

## ✨ Features
* 🔐 OAuth 2.0 authentication with Xero
* 📊 Fetches journal data from Xero API
* 🎯 Filters by tracking category (based on Google Sheet name)
* 🚫 Excludes configurable account codes
* 📅 Incremental sync (only fetches new data)
* 🔁 One-click or automated sync
* 💾 Writes to "Xero Transactions" sheet
* 🧪 Built-in System Check
* 📜 Logging for debugging
* 🔄 Retry logic for API failures
* 🧩 Modular GitHub-based script loader

---

## 🧠 Architecture

This system uses 2 files:

1. **Loader Script (PRIVATE — in Apps Script):** Holds credentials, loads code from GitHub, and handles the menu and UI.
2. **Xero Integration (PUBLIC — GitHub):** Contains all business logic and is safe to version and update centrally.

---

## ⚡ Quick Setup (10 mins)

### ✅ Checklist
* Create Xero OAuth app
* Set up Google Sheet
* Add Apps Script loader
* Install OAuth2 library
* Deploy as Web App
* Configure credentials
* Authorize + sync

### 📋 Step-by-Step Setup

#### 1️⃣ Xero Setup (2 min)
Go to: https://developer.xero.com/app/manage. Click "New app", enter any name and URL, then click "Create app". Save your **Client ID** and **Client Secret**.

#### 2️⃣ Google Sheet Setup (1 min)
Create a Google Sheet. The name must be your tracking category value exactly (Example: `WW_25_TOI`). Create a secondary sheet named `Budget, Actual, Forecast Tracking`. Set cell `B3` to your start date (e.g., `2024-01-01`).

#### 3️⃣ Apps Script Setup (5 min)
Open **Extensions → Apps Script**. Paste your `loader-template.js`. Install the OAuth2 library using Script ID: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`. Save the project.

#### 4️⃣ Deploy as Web App
Click **Deploy → New Deployment → Web App**. Set "Execute as" to "Me" and "Access" to "Anyone". Deploy and authorize.

#### 5️⃣ Configure Credentials
Get your Script ID. Create your redirect URI (`https://script.google.com/macros/d/YOUR_SCRIPT_ID/usercallback`) and add it to your Xero app settings. Update the loader script with your details:

```javascript
const PRIVATE_CONFIG = {
  CLIENT_ID: 'your_client_id',
  CLIENT_SECRET: 'your_client_secret',
  REDIRECT_URI: '[https://script.google.com/macros/d/YOUR_SCRIPT_ID/usercallback](https://script.google.com/macros/d/YOUR_SCRIPT_ID/usercallback)'
};
```


#### ▶️ Usage
##### First Time
Reload your spreadsheet. Click Xero Sync → System Check. Then click Xero Sync → 1. Authorize Xero and approve access.

##### Sync Data
Click Xero Sync → 2. Update Transactions. You’ll see a success message with the total number of transactions synced.

##### Optional: Auto Sync
Run the setupAutoSync() function once directly inside the Apps Script Editor. This will trigger the sync every hour.

#### 📊 Output
The script creates a sheet named Xero Transactions with the following columns: Date, Source ID, Account Code, Net Amount.

#### 🧪 System Check (NEW)
Accessible via Xero Sync → System Check. This tool verifies that the script is loaded, the cache exists, and checks the last sync timestamp.

#### 🔁 How Sync Works (IMPORTANT)
Incremental Sync: Stores the last sync time and only fetches new or updated data from Xero.
Safe Updates: Data is only saved to the sheet if the full sync succeeds, preventing partial data corruption.

#### ⚙️ Configuration
Edit these settings inside your main xero-integration.js file on GitHub:

JavaScript
const CONFIG = {
  SHEET_NAME: 'Xero Transactions',
  TRACKING_CATEGORY_NAME: 'Funding source',
  DATE_SHEET_NAME: 'Budget, Actual, Forecast Tracking',
  DATE_CELL: 'B3',
  EXCLUDED_ACCOUNT_CODES: ['835', '600', '610', '800', '820', '877'],
  TEST_MODE: false
};

#### 🧠 Multi-Sheet Design (KEY FEATURE)
Each Google Sheet acts as its own completely isolated data pipeline. It uses the sheet name as the specific tracking filter and maintains its own authentication and sync state.

✅ No cross-sheet interference

✅ Safe scaling across many entities

#### 📜 Logs (NEW)
You can view detailed system operations by going to Apps Script → View → Logs. This includes sync start times, API fetches, transaction counts, and detailed error messages.

#### 🚨 Troubleshooting
##### "Please authorize first"
Run: Xero Sync → 1. Authorize Xero

##### "No transactions"
Check that the sheet name is an EXACT match to the tracking value, B3 has a valid date, and the tracking category actually exists in Xero.

##### "Script cache missing"
Run: Xero Sync → Update Script from GitHub (or refresh script cache).

##### "Unauthorized"
Re-run the authorization steps from the menu.

##### Data not updating
Check the Apps Script execution logs for API failures or empty journal fetches.

### 🔒 Security Notes
CRITICAL: NEVER commit the loader script to a public repository. Keep your Client ID and Client Secret entirely private. Only the xero-integration.js code should be hosted publicly on GitHub.