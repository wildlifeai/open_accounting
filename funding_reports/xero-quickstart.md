# Quick Start Guide - Xero to Google Sheets

Get up and running in 10 minutes!

## ⚡ Quick Setup Checklist

- [ ] Create Xero OAuth app
- [ ] Set up Google Sheet with required sheet name and date cell
- [ ] Copy loader script to Apps Script
- [ ] Install OAuth2 library
- [ ] Deploy as web app
- [ ] Get Script ID and create redirect URI
- [ ] Add redirect URI to Xero app
- [ ] Update credentials in loader script
- [ ] Authorize and sync!

## 📋 Step-by-Step

### 1️⃣ Xero Setup (2 minutes)

1. Go to https://developer.xero.com/app/manage
2. Click **"New app"**
3. Fill in any name and URL
4. Click **"Create app"**
5. **Save your Client ID and Client Secret** ← Important!

### 2️⃣ Google Sheet Setup (1 minute)

1. Create or open a Google Sheet
2. **Name the file** to match your Xero tracking category (e.g., "WW_25_TOI")
3. Create a sheet called **"Budget, Actual, Forecast Tracking"**
4. Put a start date in cell **B3** (e.g., `2024-01-01`)

### 3️⃣ Apps Script Setup (5 minutes)

1. In your sheet: **Extensions** > **Apps Script**
2. Click **"+"** next to **Libraries**
3. Add library ID: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`
4. Enable the 'Show "appsscript.json" manifest file in editor appscript' option from the Settings
5. Copy code from [`appsscript-manifest.json`](funding_reports\appsscript-manifest.json) in this repo into the appscript.json
6. Copy code from [`loader-template.js`](funding_reports\loader-template.js) in this repo into the code.gs file


### 5️⃣ Configure Credentials (3 minutes)

1. In Apps Script: Click **gear icon** > Get your **Script ID**
2. Create redirect URI: `https://script.google.com/macros/d/YOUR_SCRIPT_ID/usercallback`
3. Add this URI to your Xero app settings
4. In the loader script, update:
   ```javascript
   CLIENT_ID: 'paste_your_client_id_here',
   CLIENT_SECRET: 'paste_your_client_secret_here',
   REDIRECT_URI: 'paste_your_redirect_uri_here',
   GITHUB_SCRIPT_URL: 'https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/xero-integration.js'
   ```
5. Save

### 6️⃣ Run It! (1 minute)

1. Reload your Google Sheet
2. Look for **"Xero Sync"** menu
3. Click **"Xero Sync"** > **"1. Authorize Xero"**
4. Authorize the connection
5. Click **"Xero Sync"** > **"2. Update Transactions"**
6. Done! ✅

## 🎯 Expected Results

After syncing, you'll have a new sheet called **"Xero Transactions"** with columns:
- Date
- Journal #
- Reference
- Source Type
- Source ID
- Account Code
- Account Name
- Description
- Debit
- Credit
- Net Amount
- Tax Amount
- Gross Amount
- Tracking 1
- Tracking 2

## 🚨 Common Issues

### "Menu doesn't appear"
→ Run `initialize` function manually in Apps Script, then reload sheet

### "Initialize doesn't finish running"
→ Check your Gsheet for any messages

### "Authorization failed"
→ Check Client ID, Secret, and Redirect URI match exactly

### "0 transactions found"
→ Check your sheet filename matches the Xero tracking category value

### "Can't read date from B3"
→ Make sure "Budget, Actual, Forecast Tracking" sheet exists with a date in B3

## 🔧 Customization

Want to change settings? Edit the `CONFIG` in `xero-integration.js`:

```javascript
SHEET_NAME: 'Xero Transactions',              // Output sheet name
TRACKING_CATEGORY_NAME: 'Funding source',     // Your tracking category
DATE_CELL: 'B3',                              // Start date cell
EXCLUDED_ACCOUNT_CODES: ['835', '600', '800', '820'] // Accounts to skip
```

## 📚 Need More Help?

See the full [README.md](README.md) for detailed instructions and troubleshooting.

## 🔐 Security Reminder

⚠️ **Never share or commit your loader script** - it contains your credentials!

---

Happy syncing! 🎉