# Budget sheet template

**[`Budget_sheet_template.xlsx`](Budget_sheet_template.xlsx)** — import it once and you get a
complete, correctly-named funding-source spreadsheet. The full contract, including every silent
failure mode, is in
[`dashboard/BUDGET_SHEET_TEMPLATE.md`](../dashboard/BUDGET_SHEET_TEMPLATE.md).

Four tabs:

| Tab | Role |
|---|---|
| `Funding_info` | sheet-level metadata as `key \| value` rows — owner, status, dates, contribution policy |
| `Budget` | the live baseline the dashboard reads |
| `Forecast` | optional per-quarter overrides |
| `Submitted_budget` | frozen record of what the funder was given |

It is an `.xlsx` rather than CSVs because CSV cannot carry formulas, data validation, number
formats or more than one tab. The workbook carries all four.

## Creating a new funding source

1. Upload `Budget_sheet_template.xlsx` to the right project folder in the Budgets Drive
   (`secured/` or `proposed/`).
2. Right-click it → **Open with → Google Sheets**. Drive converts it, keeping all four tab names,
   the formulas, the dropdowns and the formatting.
3. Rename the file **exactly** as its Xero *Funding source* tracking value — e.g. `SPY_27_UOA`.
   Budgets and actuals join on this string, so a mismatch shows a budget with no spend against it.
4. Delete the original `.xlsx` upload once the Sheet exists.
5. Fill in `Funding_info`, then replace the example rows on `Budget`.
6. At submission, copy the `Budget` rows onto `Submitted_budget` and never touch them again.

## What is already wired up

**Formulas.** `Contribution` is `Income − Cost` per row. `Funding_info` derives `Amount requested`,
total cost, total income, total contribution and contribution as a percentage of income straight
from the `Budget` tab, so the header can never disagree with the rows beneath it. The `Forecast` and
`Submitted_budget` tabs carry their own totals.

**Dropdowns.** `Status` (secured / proposed / archived), `Contribution policy`
(`none` / `percent_of_income:40` / `percent_of_income:20` / `per_line`), and `Project` on the
`Budget` tab.

**Colour coding.** Yellow cells are yours to fill. Grey cells are formulas — leave them alone.
Green is a header row.

## Two things that will bite

**Do not add a totals row to the `Budget` tab.** It would be read as a budget line. Totals live on
`Funding_info`, derived. (`Total` rows *are* safe on the `Forecast` tab — the parser skips them.)

**`Forecast` column headers are generated, so do not retype them.** Each one is a formula anchored
on `Funding_info!Funding start`, producing `Jul-Sep 26 Forecast` and the three quarters after it. Set
the funding start date and the headers follow.

They are deliberately **not** anchored on `TODAY()`. A `TODAY()`-based header relabels itself as time
passes while the figure beneath it stays put, so on the first day of a new quarter every forecast
silently ends up applied to a quarter it was never meant for. Anchoring on the grant's own start date
keeps each label fixed for the life of the grant.

The format must remain exactly `MMM-MMM YY Forecast`; the year is not optional, and
`Jul-Sep Forecast` fails to parse, discarding that whole quarter without an error. If you need more
than four quarters, copy a header cell sideways and add 3 to both `EOMONTH` offsets.

Forecast rows must start with the `Xero Inventory Item` in `CODE - Name` form
(`SPY_26_EXAMPLE_001 - Baseline model assessment`) — readable *and* resolvable, since the dashboard
splits on `' - '` and keys on the code.

## Filling in the Budget tab

**One row per milestone, or several rows sharing one `Xero Inventory Item`.** The example splits
each milestone into a Data Scientist row and a Project Manager row; both carry the same item code,
so they roll up as one milestone while staying legible. Do **not** itemise by Xero account — there
is no `*Account` column, because nothing reads it and predicting the account split months ahead is
wasted effort.

**`Milestone` is the load-bearing column, `Description` is not.** `Milestone` is the grouping key
for the Overview breakdown, the row label in the tracking grid, and a filter dimension. Leave it
blank and the whole funding source collapses into one `(unassigned)` group. `Description` is read by
nothing — use it for the human detail, as the example does with roles.

**Dates** — real dates formatted `DD/MMM/YY`. A line whose End precedes its Start is dropped
entirely, and two-digit years are the usual cause: `30/Jun/01` is the year 2001.

## Checking your work

Open the Funding Cockpit and look at the **Health** panel. It names the sheet, the row and what to
do. Importing this template unaltered produces no findings, so anything reported afterwards is your
own data. If a sheet you have just created does not appear at all, the usual causes are a tab not
named `Budget`, a missing `Start`/`End`/`Cost` column, or a filename that does not match the Xero
tracking value.
