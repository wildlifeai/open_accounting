# Funding-source budget sheet — template and contract

> **Creating or updating a sheet?** Start from the importable CSVs in
> [`budget_templates/`](../budget_templates/) rather than transcribing the tables below —
> `Budget.csv` and `Forecast.csv` import straight into Google Sheets and land with the correct tab
> names. This document is the reference for what every field means and how each one fails.

Every funding source has one Google Sheet in the Budgets Drive, named **exactly** as its Xero
*Funding source* tracking value (e.g. `WW_25_TOI`), living in its project's `secured/` or
`proposed/` folder.

**At most three tabs:**

| Tab | Role |
|---|---|
| `Budget` | the live baseline the cockpit reads. Change only for a genuine re-budget. |
| `Forecast` | per-quarter overrides, where you know something the budget does not. Optional. |
| `Submitted_budget` | frozen record of what the funder was actually given. Never edited. |

Nothing else. Actuals come live from Xero, so a per-sheet copy of them is a second version of the
truth and will drift.

---

## `Budget` tab layout

A metadata block, one blank row, then the column header row, then the budget lines. Keys go in
column **A**, values in column **B**.

```
A                        | B                      | C ...
-------------------------|------------------------|--------
Funding source           | WW_25_TOI              |
Project                  | Wildlife Watcher       |
Funder                   | Toi Foundation         |
Status                   | secured                |
Funding start            | 10/Aug/25              |
Funding end              | 10/Dec/26              |
Contribution policy      | percent_of_income:40   |
Amount requested         | 143436                 |
Amount secured           | 143436                 |
Decision date            | 2025-08-01             |
Owner                    | victor@wildlife.ai     |
Last reviewed            | 2026-08-11             |
Notes                    | Core roles to Dec 26   |
                         |                        |
Description | Start | End | Cost | Income | Contribution | *Account | Milestone | Xero Inventory Item | Project | Comments
General management 0.2 FTE | 03/Nov/25 | 02/Aug/26 | 23101 | 0 | 0 | Salaries (477) | General management | WW_25_TOI_002 | General | 2.3k monthly for GM
Product management 0.4 FTE | 05/Jan/26 | 10/Jan/27 | 53571 | 0 | 0 | Salaries (477) | Product management | WW_25_TOI_003 |  |
Toi grant income           | 10/Aug/25 | 10/Dec/26 | 0 | 143436 | 57374 | Project Contract Income (181) | Grant income | WW_25_TOI_006 |  | 40% contributes to General
```

## Metadata fields

| Key | Required | Meaning |
|---|---|---|
| `Funding source` | yes | Must equal the file name **and** the Xero *Funding source* tracking value. The cockpit joins budgets to actuals on this string. |
| `Project` | yes | Default project for lines with a blank `Project` cell. Must be a Xero *Projects* tracking value. |
| `Funder` | yes | The organisation. Free text; for reporting and for spotting two applications to the same funder. |
| `Status` | yes | `secured` \| `proposed` \| `archived`. Must agree with the folder the sheet sits in — disagreement is a health-check failure, not a preference. |
| `Funding start` / `Funding end` | yes | The grant period. `Funding end` is what lets the cockpit flag actuals still landing against a finished grant. |
| `Contribution policy` | yes | How much of this source's income funds `General`. One of `none`, `percent_of_income:<n>`, `per_line`. See below. |
| `Amount requested` | yes | GST-exclusive total sought. |
| `Amount secured` | proposed: `0` | GST-exclusive total confirmed. |
| `Decision date` | proposed only | When the funder decides. Blank for secured. Funders ask for this on application forms. |
| `Owner` | yes | Email of whoever maintains this sheet. Health checks are addressed to this person. |
| `Last reviewed` | yes | Date last checked against reality. Staleness is otherwise invisible. |
| `Notes` | no | Free text. Not parsed — never put a number here that something else needs. |

### `Contribution policy`

This replaces the current practice of hand-computing overheads each quarter and writing the split
as prose into a Comments cell, which nothing can read and which drifts from the project budgets.

| Value | Meaning |
|---|---|
| `none` | Project-specific grant that disallows overheads. Contributes nothing to General. |
| `percent_of_income:40` | 40% of this source's income funds General. The cockpit derives the amount. |
| `per_line` | The split is expressed per budget line via the `Project` column — used where some lines are General work and others are project work (e.g. `WW_25_TOI`). |

Once this field is populated, the mirrored negative-cost "Overheads from projects" lines in the
General budget should be deleted. They are a hand-maintained duplicate of a derivable number.

## Column definitions

| Column | Required | Notes |
|---|---|---|
| `Description` | no | Human label. |
| `Start`, `End` | **yes** | A line with either unparseable is **silently skipped**. Use `DD/MMM/YY`. Check the year: `30/Jun/01` parses as 2001. |
| `Cost` | **yes** | GST-exclusive. Use `0`, not blank. |
| `Income` | yes | GST-exclusive. A line where `Cost` and `Income` are both `0` is **silently skipped**. |
| `Contribution` | yes | `Income − Cost`; the margin funding General. Already inside `Income`, so never counted twice. |
| `*Account` | yes | Xero chart-of-accounts label exactly as `Name (code)`, e.g. `Salaries (477)`. **Per line, not per milestone** — see below. |
| `Milestone` | yes | Human milestone name. |
| `Xero Inventory Item` | yes | The `{SOURCE}_{NNN}` product/service code, e.g. `WW_25_TOI_002`. This is the milestone dimension; without it a line falls out of the quarterly tracking grid. |
| `Project` | yes (column) | Per-line override. Blank means "use the `Project` metadata value". The column must exist even if every cell is blank. |
| `Comments` | no | Not parsed. |

Two columns are marked required-by-contract rather than required-by-parser: `*Account` and
`Xero Inventory Item` are technically optional to `BudgetReader`, but account-level P&L and
milestone tracking both break without them. Treat them as mandatory.

Be aware `*Account` has **no consumer in the code today** — the only reads of `.account` anywhere
are in a `WebApp.js` diagnostic that logs *actuals*, not budgets. It is captured for the quarterly
account-level P&L the board needs, which is not built yet, and for reconciliation against Xero
actuals, which do carry accounts. That is why health check A5 is `info` rather than a warning:
nothing on screen is wrong without it.

## One milestone spans many accounts

A milestone is a chunk of work; an account is what kind of cost it is. They are separate dimensions
and a budget row is their intersection, so **never create a milestone per account** — repeat the
same `Xero Inventory Item` across as many rows as the milestone needs.

Rollup is keyed on `project||source||milestone` (`Aggregator.js:29`), and account appears in no key,
so rows sharing an item code aggregate into one milestone with the accounts preserved underneath.
This is already how the live sheets work: `WW_25_TOI_002` spans eight accounts as a single
milestone, `WW_25_TOI_003` spans four.

A data-scientist milestone is therefore three rows — the contractor on `Contractors (410)`,
recruitment on `Advertising (400)`, software on `Subscriptions (485)` — all carrying the same item
code. Splitting it into three milestones would triple the rows in the quarterly tracking grid for
no gain.

## `Forecast` tab

Optional, and only worth creating when you need to override the budget. **A quarter with no entry
here falls back to the budget baseline**, so an empty or absent `Forecast` tab is a valid state —
it means "the budget is still our best estimate". Record a forecast when you know something the
budget does not: a delayed hire, a grant ending early, a re-profiled milestone.

Layout, as parsed by `BudgetReader.parseForecastTab_`:

```
A                              | B                     | C                     | D
-------------------------------|-----------------------|-----------------------|----------
Revenue                        | Jul-Sep 26 Forecast   | Oct-Dec 26 Forecast   | Comments
WW_25_TOI_006 - Grant income   | 35000                 | 35000                 |
                               |                       |                       |
Expenses                       | Jul-Sep 26 Forecast   | Oct-Dec 26 Forecast   | Comments
WW_25_TOI_002 - General mgmt   | 7108                  | 0                     | GM role vacant from Oct
WW_25_TOI_003 - Admin & Comms  | 2390                  | 0                     | comms advisor departed
                               |                       |                       |
Funding Source Details         |                       |                       |
```

Rules that are easy to get wrong:

- Section headers in column A must read exactly `Revenue` or `Expenses`. Rows before the first
  section header are ignored.
- **Quarter column headers must match `MMM-MMM YY Forecast` exactly** — e.g. `Jul-Sep 26 Forecast`.
  Anything else is silently ignored, so a typo in one heading quietly discards that quarter's
  forecast with no error. Health check **A7** exists for this.
- Column A rows under a section must be `CODE - Name`, where `CODE` is the `Xero Inventory Item`
  from the `Budget` tab. A row whose code cannot be parsed is skipped.
- Blank rows and rows whose column A starts with `Total` are skipped.
- Parsing **stops entirely** at a column-A value of `Funding Source Details`. Anything below that
  line is invisible, which makes it a useful place for working notes.
- Each section can carry its own quarter columns; they are re-read per section header.
- A `Comments` column (header exactly `Comments`) attaches one free-text note per milestone,
  surfaced in the tracking grid. Use it to say *why* a forecast differs from the budget.

Forecast values are per quarter, not per month, and are absolute amounts rather than adjustments.

## Common failures

| Symptom | Cause |
|---|---|
| Budget missing from the dashboard entirely | no `Budget` tab, or a required column absent — the whole file is skipped with only a log line |
| Some lines missing | `Cost` and `Income` both `0`, or an unparseable/reversed date |
| Everything in one `(unassigned)` row | `Xero Inventory Item` blank |
| Actuals present, budget shows `0` | sheet name does not match the Xero *Funding source* value |
| Spend looks impossibly low | Xero transactions missing the `Projects` tracking tag are dropped entirely |

## Parser change this template requires

`BudgetReader.parseBudgetFile_` currently treats **row 1** as the column header row
(`const header = data[0]`). The metadata block above the columns therefore needs the parser to
**locate** the header row rather than assume it — scan the first ~30 rows for one containing both
`Start` and `Cost`, then read metadata as key/value pairs from column A/B of the rows above it.

Deploy that change **before** rolling the template out to existing sheets, or every sheet with a
metadata block will be skipped for a missing column.
