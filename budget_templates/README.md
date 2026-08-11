# Budget sheet templates

Importable CSV starting points for a funding-source spreadsheet. The full contract — every
column, every field, and the silent failure modes — is in
[`dashboard/BUDGET_SHEET_TEMPLATE.md`](../dashboard/BUDGET_SHEET_TEMPLATE.md). These files are
the version you can actually import.

| File | Becomes the tab |
|---|---|
| `Budget.csv` | `Budget` — the live baseline the cockpit reads |
| `Forecast.csv` | `Forecast` — per-quarter overrides, optional |

**The filenames matter.** Google Sheets names an imported sheet after the file, so importing
`Budget.csv` gives you a tab called `Budget` — which is the name the cockpit looks for. Check it
afterwards anyway: if a tab of that name already exists, Sheets appends a number and the cockpit
will not find it.

## Creating a new funding source

1. In the Budgets Drive, open the right project folder, then `secured/` or `proposed/`.
2. **New → Google Sheets**, and name the file **exactly** as its Xero *Funding source* tracking
   value — e.g. `WAI_27_OMV`. Budgets and actuals join on this string, so a mismatch means the
   sheet shows a budget with no spend against it.
3. **File → Import → Upload** `Budget.csv`, and choose **Insert new sheet(s)**.
4. Repeat for `Forecast.csv` if you need overrides. Skip it otherwise — a quarter with no entry
   falls back to the budget baseline, so an absent `Forecast` tab is a perfectly good state.
5. Delete the default `Sheet1`.
6. Fill in the metadata block, then replace the example lines.
7. At submission, duplicate the `Budget` tab (right-click → **Duplicate**) and rename the copy
   `Submitted_budget`. That is the frozen record of what the funder was given; never edit it.

No other tabs. The dashboard flags extras (health check A3), because actuals live in Xero and a
per-sheet copy of them drifts.

## Filling it in

**Metadata block** — keys in column A, values in column B, above the column header row. All
thirteen fields are expected; `Amount secured` is `0` while a bid is proposed, and
`Decision date` is blank once it is secured.

`Contribution policy` is the one people skip, and it is the field that finally replaces
hand-computing overheads each quarter. One of:

- `none` — project-specific grant that disallows overheads
- `percent_of_income:40` — 40% of this source's income funds General
- `per_line` — the split is expressed per line via the `Project` column

**Dates** — `DD/MMM/YY`, e.g. `01/Oct/26`. Check the year: `30/Jun/01` parses as **2001**, and a
line whose End precedes its Start is dropped entirely (health check B3).

**Amounts** — plain numbers. `24000`, not `$24,000`. A currency symbol is tolerated but a comma
inside an unquoted CSV field will split the row.

**`Contribution`** must equal `Income − Cost` on every line, or health check B5 fires. In the
example the three lines sum to $36,000 cost against $60,000 income, giving $24,000 of
contribution — exactly the 40% the metadata declares.

**`Xero Inventory Item`** is the milestone code, formatted `SOURCE_YY_NAME_NNN` — e.g.
`WAI_27_EXAMPLE_001`. Without it a line falls outside quarterly tracking (B4), and on a budget
that is mostly salaries that means most of the money.

### One milestone, many accounts

**Never create a milestone per Xero account.** A milestone is a chunk of work; an account is what
kind of cost it is. They are different dimensions, and a budget row is the intersection of the two,
so a milestone spans as many accounts as it needs — just repeat the same `Xero Inventory Item` on
each row.

The example shows it: `WAI_27_EXAMPLE_001` is one milestone across three rows —

| Description | `*Account` | `Xero Inventory Item` |
|---|---|---|
| Data scientist contractor 0.4 FTE | `Contractors (410)` | `WAI_27_EXAMPLE_001` |
| Recruitment for the data scientist | `Advertising (400)` | `WAI_27_EXAMPLE_001` |
| Analysis software subscriptions | `Subscriptions (485)` | `WAI_27_EXAMPLE_001` |

The dashboard rolls up on project + funding source + milestone, so those three rows appear as one
milestone with a combined budget, and the accounts stay available underneath for reconciliation
against Xero. Your existing sheets already work this way: `WW_25_TOI_002` spans eight accounts —
salaries, rent, insurance, accounting, advertising, contractors, general expenses and volunteer
expenses — as a single milestone.

Splitting one milestone into three because the money lands in three accounts would triple the rows
in the tracking grid and make quarterly reporting unreadable, for no gain.

**`Project`** — leave blank to use the sheet's `Project` metadata value. Fill it only to send a
line elsewhere, as the example's general-management line does to `General`.

## The Forecast tab

Only worth creating when you know something the budget does not: a delayed hire, a grant ending
early, a re-profiled milestone. Two rules break it silently if you get them wrong:

- Quarter column headers must read **exactly** `MMM-MMM YY Forecast`, e.g. `Jul-Sep 26 Forecast`.
  Anything else is ignored, discarding that whole quarter with no error (A7).
- Rows must be `CODE - Name`, where `CODE` matches a `Xero Inventory Item` on the `Budget` tab
  (A8, A9).

Parsing stops at a column-A value of `Funding Source Details`, which makes everything below it a
safe place for working notes.

## Checking your work

Open the Funding Cockpit and look at the **Health** panel. It names the sheet, the row and what to
do — reversed dates, missing item codes, unexpected tabs, unparseable forecast headers. If a sheet
you have just created does not appear at all, the usual causes are a tab not named `Budget`, a
missing `Start`/`End`/`Cost` column, or a filename that does not match the Xero tracking value.
