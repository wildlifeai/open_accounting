# Project Reports & Quarterly Budgets

Apps Script sources for generating project reports and aggregating budgets for Xero.

Both scripts here are deployed with [clasp](https://github.com/google/clasp). They are
**not** fetched from GitHub at runtime — see [Why there is no loader](#why-there-is-no-loader).

## Contents

| File | Apps Script project |
|---|---|
| `create_quarterly_budgets.js` | standalone project (quarterly budgets) |
| `funding-aggregator.js` | container-bound to the `PROJECT_overview` spreadsheet |
| `appsscript.json` | manifest: OAuth2 library + OAuth scopes |

These belong to two different Apps Script projects, so one `.clasp.json` cannot push both.
Copy `.clasp.json.example` to `.clasp.json` with the `scriptId` of whichever project you are
deploying, or give each script its own subfolder if you want to push both from here.

## `create_quarterly_budgets.js`

Parses the funding-source budget sheets in each project's `secured` folder and produces a
summary spreadsheet plus Xero import CSVs.

### Configuration

- **Budgets Folder ID**: `10105co6S5qHFSVVg0pb0fkoPidN3ScJZ`
- **Projects**: Spyfish Aotearoa (SPY), General (GEN), Wild About AI (WAI), Wildlife Watcher (WW).
  These abbreviations match the funding-source file names already in Drive (`WAI_25_OMV`,
  `WW_25_TOI`, `SPY_26_UOA`) and the Xero *Funding source* tracking values. Do not rename them
  without renaming those too.
- **Target Sheet**: the tab named `Budget, Actual, Forecast Tracking`.
- **Extracted Columns**: G, I, J, K, L (indices 6, 8, 9, 10, 11).
- **Chart of accounts**: first sheet of `1VHtZsZRzJJ29tt3SRDXnIP6ebKChoDtlHqyV5Ua-3Yg`, column J.

### Logic

1. **Date parsing** — April 1 financial-year start. A date in row 4 (e.g. `30-Jun-2026`) maps to
   an FY quarter (`Q1`) and its amounts land in the middle month of that quarter (`May 2026`).
2. **Account validation** — only accounts present in the master chart of accounts are kept.
3. **Summary sheet** — creates or updates `YYQX_PROJECT` (e.g. `26Q2_SPY`) in the project folder.
4. **Funding-source tabs** — one tab per funding source, plus `OVERALL_PROJECT`, with column A
   prepopulated with valid accounts and a column per month. Tab names are sanitised and
   truncated to 31 characters to satisfy the Sheets limit.
5. **CSV export** — writes `<budget>_XERO_IMPORT.csv` per funding source, RFC 4180 quoted.
6. **Diff** — reads existing budgets from the Xero API and logs a checklist of which CSVs
   actually need uploading. Files that fail to open are skipped and listed at the end rather
   than aborting the run.

### Xero Budgets API

The Xero API supports **GET only** for Budgets — there is no POST/PUT endpoint. So the script
generates CSVs for manual import and uses the API purely to diff against what is already in
Xero, so you only upload the budgets that changed.

### Setup

1. Add the OAuth2 library as `OAuth2` — already declared in `appsscript.json`, so `clasp push`
   carries it. If setting up by hand: Libraries → **+** → Script ID
   `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF` → latest version.
2. **Project Settings → Script Properties**, add:

   | Property | Value |
   |---|---|
   | `XERO_CLIENT_ID` | from your Xero app |
   | `XERO_CLIENT_SECRET` | from your Xero app |

   Never put these in source. The script reads them via `getSecret_()`.
3. Run `logXeroRedirectUri` and register the logged URI in the Xero app's **Redirect URIs**.
4. Run `showAuthorizationUrl`, open the URL, approve the organisation.
5. Run `generateAndUploadQuarterlyBudgets`.

Entry points are the only functions without a trailing underscore: `logXeroRedirectUri`,
`showAuthorizationUrl`, `authCallback`, `generateAndUploadQuarterlyBudgets`.

## `funding-aggregator.js`

Bound to the `PROJECT_overview` spreadsheet. Adds a **Funding Tools → Update Funding Data**
menu that aggregates the `secured` and `proposed` funding sheets sitting alongside it into the
`Overview` tab.

## Why there is no loader

Earlier versions used `loader-budgets-template.js` and `loader_template.js`, which fetched these
files from a raw GitHub URL and executed them with `new Function` / `eval`. Both have been
removed. That pattern meant:

- any change to the referenced branch executed with the full permissions of whoever ran the
  script, and the budgets loader passed the Xero client secret straight into the fetched code;
- the fetched code was cached in Script Properties, so it kept running even after the remote
  was cleaned up;
- neither loader checked the HTTP status, so a 404 body was passed to the interpreter — which is
  exactly what happened when the `feature/budget_xero` branch was deleted after merging.

The loaders existed to get updated code into Apps Script without copy-paste. `clasp push` does
that properly, so they are no longer needed.

**If you previously installed the budgets loader**, delete the `BUDGETS_SCRIPT_CACHE` property
from that project's Script Properties — cleaning the repo does not clear the cache.

## Deploying

```bash
cp .clasp.json.example .clasp.json   # then set scriptId
clasp push
```

Pull before editing and push after, so the repo stays the source of truth. `clasp push` replaces
the remote project with the local files, so a stale checkout will silently revert work done in
the browser editor.
