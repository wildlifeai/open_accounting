# Funding-source budget sheet — template and contract

Every funding source has one Google Sheet in the Budgets Drive, named **exactly** as its Xero
*Funding source* tracking value (e.g. `WW_25_TOI`), living in its project's `secured/` or
`proposed/` folder.

**At most two tabs:**

| Tab | Role |
|---|---|
| `Budget` | the live baseline the cockpit reads. Change only for a genuine re-budget. |
| `Submitted_budget` | frozen record of what the funder was actually given. Never edited. |

Nothing else. Actuals come live from Xero; forecasts live in the central Cockpit Forecast sheet.
A per-sheet copy of either is a second version of the truth and will drift.

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
| `*Account` | yes | Xero chart-of-accounts label exactly as `Name (code)`, e.g. `Salaries (477)`. |
| `Milestone` | yes | Human milestone name. |
| `Xero Inventory Item` | yes | The `{SOURCE}_{NNN}` product/service code, e.g. `WW_25_TOI_002`. This is the milestone dimension; without it a line falls out of the quarterly tracking grid. |
| `Project` | yes (column) | Per-line override. Blank means "use the `Project` metadata value". The column must exist even if every cell is blank. |
| `Comments` | no | Not parsed. |

Two columns are marked required-by-contract rather than required-by-parser: `*Account` and
`Xero Inventory Item` are technically optional to `BudgetReader`, but account-level P&L and
milestone tracking both break without them. Treat them as mandatory.

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
