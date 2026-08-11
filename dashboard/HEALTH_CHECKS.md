# Self-policing and system health checks

The cockpit's numbers are only as good as the budgets in Drive and the coding in Xero, and both
fail **silently** today. `BudgetReader` skips a malformed line with a `Logger.log` nobody reads; an
untagged Xero transaction vanishes from every total without a trace. The purpose of health checks is
to convert silent wrongness into a named, owned, actionable item that the GM or a project lead can
fix themselves.

Design principle: **every finding names what is wrong, where, who owns it, and what to do.** A
check that only says "something looks off" creates work instead of removing it.

## Where it plugs in

There is already a mechanism: `Aggregator.buildSnapshot()` produces `dataFlags`, and
`JavaScript.html` renders them into `#flags`. Extend that rather than inventing a parallel path.

```
BudgetReader ──┐
XeroClient   ──┼──► Aggregator.buildSnapshot() ──► snapshot.health[] ──► Health panel
metadata     ──┘                                                        (grouped by owner)
```

Findings belong in the snapshot so they are cached with everything else and cost nothing to
display. `snapshot.dataFlags` becomes the legacy alias for `health` filtered to `severity=error`.

## Finding shape

```js
{
  id: 'B3',                            // stable check id, so a finding can be documented
  severity: 'error',                   // error | warning | info
  category: 'Data quality',
  title: 'Budget line ends before it starts',
  detail: 'WW_25_TOI line 24: End 30/Jun/01 precedes Start 01/Feb/26.',
  fundingSource: 'WW_25_TOI',
  project: 'Wildlife Watcher',
  owner: 'victor@wildlife.ai',         // from the sheet's metadata block
  amount: 2000,                        // value at risk, when quantifiable
  action: 'Fix the End date on that line. 30/Jun/01 parses as year 2001.',
  link: 'https://docs.google.com/spreadsheets/d/…'
}
```

`severity` means something specific:

| | Meaning |
|---|---|
| **error** | A number on the dashboard is currently wrong. |
| **warning** | A number may be wrong, or will be soon. |
| **info** | Hygiene. Nothing is wrong yet. |

`owner` comes from the `Owner` key in the sheet's metadata block (see
[BUDGET_SHEET_TEMPLATE.md](BUDGET_SHEET_TEMPLATE.md)), which is what makes per-lead filtering
possible. It pairs naturally with the existing `PERMISSIONS` layer: a project lead sees findings for
their projects, the GM sees everything.

## Check catalogue

### A. Sheet structure

| id | Sev | Check | Action shown |
|---|---|---|---|
| A1 | error | Funding-source file has no `Budget` tab | Whole file is invisible. Add or rename the tab. |
| A2 | error | `Budget` tab missing `Start`, `End` or `Cost` | Whole file is invisible. Add the column. |
| A3 | warning | Tabs beyond `Budget` / `Forecast` / `Submitted_budget` | Retire the extra tab; actuals live in Xero. |
| A4 | warning | No `Project` column | Lines cannot be split across projects. |
| A5 | warning | No `*Account` column | Account-level P&L impossible for this source. |
| A6 | info | No `Submitted_budget` tab | No frozen record of what the funder was given. |
| A7 | error | `Forecast` tab column header does not match `MMM-MMM YY Forecast` | That quarter's forecast is silently discarded. Name the column exactly, e.g. `Jul-Sep 26 Forecast`. |
| A8 | warning | `Forecast` tab row whose column A is not a parseable `CODE - Name` | That row is skipped. Match the `Xero Inventory Item` from the `Budget` tab. |
| A9 | info | `Forecast` entry for an item code absent from the `Budget` tab | Forecasting a milestone that no longer exists. |

A missing or empty `Forecast` tab is **not** a finding — a quarter with no override falls back to
the budget baseline, so an absent tab legitimately means "the budget is still our best estimate".

### B. Data quality

| id | Sev | Check | Action shown |
|---|---|---|---|
| B1 | warning | Lines skipped because `Cost` and `Income` are both 0 | Report count. Use 0 deliberately or delete the row. |
| B2 | error | Unparseable `Start` or `End` | Line is invisible. Use `DD/MMM/YY`. |
| B3 | error | `End` before `Start` | Check the year — `30/Jun/01` parses as 2001. |
| B4 | error | Blank `Xero Inventory Item` | Line falls out of the tracking grid. Report count **and share of budget value**. |
| B5 | warning | `Contribution` ≠ `Income − Cost` | Recompute or explain. |
| B6 | warning | `*Account` not in the Xero chart of accounts | Fix the label to `Name (code)` exactly. |

### C. Metadata

| id | Sev | Check | Action shown |
|---|---|---|---|
| C1 | error | Required metadata key missing | Name the key. |
| C2 | error | `Status` disagrees with the folder | Move the file or fix the value — the cockpit trusts the folder for forecasts. |
| C3 | error | `Funding source` ≠ file name | Budgets and actuals will not join. |
| C4 | warning | `Last reviewed` older than 90 days | Review it, then update the date. |
| C5 | warning | `Funding end` in the past, sheet still in `secured/` | Archive it, or extend the end date. |
| C6 | error | `Contribution policy` missing or unparseable | General's income cannot be derived. |
| C7 | warning | `proposed` with no `Decision date` | Needed for pipeline forecasting and funder forms. |

### D. Xero coding — aimed at the bookkeeper

| id | Sev | Check | Action shown |
|---|---|---|---|
| D1 | error | Actual lines with no `Projects` tag | **Dropped from every total.** Report count and value; link the transactions. |
| D2 | warning | Actual lines with no `Funding source` tag | Land in `(unassigned)`. |
| D3 | warning | Actual lines with no item code | Outside the tracking grid. |
| D4 | error | Actuals coded to a `Funding source` with no budget sheet | Either the sheet is missing or the tag is a typo. |
| D5 | warning | Budget sheet with zero actuals though its period has started | Nothing is being coded to it. |
| D6 | error | Actuals against a funding source whose `Funding end` has passed, or whose sheet is archived | Almost always a stale recurring journal or template. |
| D7 | error | Spend counted on an excluded balance-sheet account | Only reachable once `EXCLUDED_ACCOUNTS` is wired up. |

### E. Reconciliation

| id | Sev | Check | Action shown |
|---|---|---|---|
| E1 | warning | Actuals exceed budget for a funding source | Re-budget or explain to the funder. |
| E2 | warning | Under-spend risk: proportion spent well below proportion of period elapsed | Funders care about underspend as much as overspend. |
| E3 | error | Same item code in two funding sources | One cost billed twice. |
| E4 | warning | Same `*Account` + `Description` in two sources with overlapping dates | Double-funding signal — e.g. the same FTE in two grants. |
| E5 | warning | Salary actuals attributed differently from budgeted salary lines | **The payroll-template drift detector.** Nothing in Xero reports a stale repeating journal; this is the only way it surfaces. |
| E6 | info | Residual hand-entered overhead lines alongside a derived `Contribution policy` | Delete the manual duplicate. |

### F. System health — aimed at the GM and maintainer

| id | Sev | Check | Action shown |
|---|---|---|---|
| F1 | error | Xero not connected or token invalid | Run the reconnect step; actuals are stale meanwhile. |
| F2 | warning | Snapshot older than 2× the refresh interval | The refresh trigger may be broken. |
| F3 | error | Required Script Properties missing | Name which. |
| F4 | warning | A Xero scope needed by a feature in use is absent | Re-consent required. |
| F5 | info | Last refresh time, duration, sheets read, lines parsed | Trend tells you when the 6-minute limit is approaching. |

## UI

A **Health** panel on the Overview tab, collapsed by default, showing counts by severity
(`3 errors · 11 warnings · 6 info`) and expanding to a list grouped by category. Each row shows
title, where, the value at risk if any, and the action. Errors also surface as a persistent banner,
because an error means a number on screen is wrong right now.

Filter controls: **mine / all** (by `owner`), and by project. Project leads land on *mine*.

Do not sort purely by severity — sort by **value at risk** within severity, so a $40,000
misattribution outranks a missing description.

## Implementation notes

* Put the checks in a new `HealthCheck.js`, pure functions over `(budgets, actualLines, metadata)`
  returning findings. Keep it free of `DriveApp`/`UrlFetchApp` so `Tests.js` can cover it offline —
  it is the first part of this codebase that is genuinely unit-testable, so take the opportunity.
* `BudgetReader` must **report** what it currently discards. Today it `continue`s past bad rows;
  it needs to collect `{file, row, reason}` and return them alongside the lines. This is the single
  biggest change, and most of A and B depend on it.
* D1–D3 need the count **and** the summed value of affected lines; a count alone doesn't convey
  whether it matters.
* Cap each check's findings (say 20) and report the overflow count — never truncate silently, which
  is the failure mode the checks exist to prevent.
* Checks must never throw. A failing check reports itself as an `info` finding and the rest proceed.

## Sequencing

Two things gate this work:

1. **`sync-from-apps-script` must merge first.** It holds ~1,030 lines of `Aggregator.js`,
   `WebApp.js` and `JavaScript.html` changes. Building the health panel on a third branch before it
   lands guarantees conflicts in exactly those files.
2. **The metadata template drives categories C and much of D and E.** Those checks are
   unimplementable until sheets carry `Owner`, `Status`, `Funding end` and `Contribution policy`.

So: merge the sync branch, then implement the `BudgetReader` reporting change plus categories A, B
and F — which need no new sheet data and would surface real problems today. Categories C, D6, E4 and
E5 follow as the metadata block rolls out.
