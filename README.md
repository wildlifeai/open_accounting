# open_accounting

Scripts and documentation related to the budgeting and accounting frameworks of Wildlife.ai.

These are Google Apps Script projects that join **budgets** (Google Sheets in Drive) to
**actuals** (Xero), so the board, the GM, project leads and funders all work from the same
numbers.

## Documentation

| | |
|---|---|
| **Start here** | [`AGENTS.md`](AGENTS.md) — quickstart, commands, non-negotiables, repo map |
| **Deep guide** | [`.agents/skills/SKILL.md`](.agents/skills/SKILL.md) — invariants, the Xero data model, known defects, and the traps that have actually bitten |
| Funding Cockpit dashboard | [`dashboard/README.md`](dashboard/README.md) |
| — for non-technical users | [`dashboard/GM_GUIDE.md`](dashboard/GM_GUIDE.md) |
| — effect on budget procedures | [`dashboard/BUDGET_PROCEDURES_ADDENDUM.md`](dashboard/BUDGET_PROCEDURES_ADDENDUM.md) |
| Quarterly budget generation | [`project_reports/README.md`](project_reports/README.md) |
| Funding reports (Xero) | [`funding_reports/xero-quickstart.md`](funding_reports/xero-quickstart.md) |

`CLAUDE.md` is just `@AGENTS.md`, so AI assistants and developers read the same guide.

> **Before pushing to Apps Script, read the Deployment Invariant in `SKILL.md`.** The deployed
> projects can be edited in the browser and have drifted ahead of this repo before;
> `clasp push` replaces the remote and deletes anything not present locally.
