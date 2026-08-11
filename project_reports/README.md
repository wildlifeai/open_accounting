# Project Reports

Apps Script source for aggregating funding data into the `PROJECT_overview` spreadsheet.

Deployed with [clasp](https://github.com/google/clasp). Nothing here is fetched from GitHub at
runtime — see [Why there is no loader](#why-there-is-no-loader).

## `funding-aggregator.js`

Container-bound to the `PROJECT_overview` spreadsheet. Adds a **Funding Tools → Update Funding
Data** menu that aggregates the `secured` and `proposed` funding sheets sitting alongside it into
the `Overview` tab.

Because it is container-bound, it does not appear in `clasp list-scripts`. To manage it with clasp
you need its script id from the container: in the spreadsheet, **Extensions → Apps Script**, then
**Project Settings → IDs**.

`appsscript.json` carries the manifest (OAuth2 library and scopes). Copy `.clasp.json.example` to
`.clasp.json` and set `scriptId` before pushing.

## Retired

**`create_quarterly_budgets.js`** generated `YYQX_PROJECT` summary spreadsheets and Xero budget
import CSVs, and diffed them against the Xero Budgets endpoint. Deleted on 2026-08-11 along with
its standalone `quarterly_budgets` Apps Script project, because:

- the summary sheets and monthly aggregation are superseded by the Funding Cockpit's breakdown and
  quarterly tracking;
- nobody uses Xero's budget-variance reports, so there is no consumer for the CSVs;
- a Xero budget is keyed on account and period only — the funding source survives merely as text in
  the budget's name and the milestone dimension is lost — so it cannot represent the four-dimension
  model the cockpit is built on;
- it required the `Budget, Actual, Forecast Tracking` tab, which the three-tab rule removes.

The code is in git history if any of it is ever wanted. See
[`.agents/skills/SKILL.md`](../.agents/skills/SKILL.md) §8.

**`loader-budgets-template.js`** and **`loader_template.js`** — see below.

## Why there is no loader

Earlier versions fetched these files from a raw GitHub URL and executed them with `new Function` /
`eval`. Both loaders have been removed. That pattern meant:

- any change to the referenced branch executed with the full permissions of whoever ran the
  script, and the budgets loader passed the Xero client secret straight into the fetched code;
- the fetched code was cached in Script Properties, so it kept running even after the remote was
  cleaned up;
- neither loader checked the HTTP status, so a 404 body was passed to the interpreter — which is
  exactly what happened when the `feature/budget_xero` branch was deleted after merging PR #6.

The loaders existed to get updated code into Apps Script without copy-paste. `clasp push` does that
properly, so they are no longer needed.

A third loader still exists in [`../funding_reports/`](../funding_reports/) and needs the same
treatment.

## Deploying

```bash
cp .clasp.json.example .clasp.json   # then set scriptId
clasp push
```

Pull before editing and push after, so the repo stays the source of truth. `clasp push` replaces
the remote project with the local files, so a stale checkout will silently revert work done in the
browser editor.
