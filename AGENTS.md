# Agent guide — Wildlife.ai accounting

**Google Apps Script** tooling that joins Wildlife.ai's budgets (Google Sheets in Drive) to its
actuals (Xero) so the GM, board, project leads and funders can each see the same numbers. The
flagship is the **Funding Cockpit** dashboard; alongside it sit quarterly-budget generation,
funding reports and chart-of-accounts helpers.

This repo is **not** the running system. Every script here executes as an Apps Script project in
the cloud, and those projects can be edited in the browser. Treating the repo as the sole source
of truth has already cost real work.

**Before changing any script, pushing to Apps Script, or touching Xero, read
[`.agents/skills/SKILL.md`](.agents/skills/SKILL.md)** — the invariants and the traps that have
actually bitten here. This file is only the quickstart.

## Setup and everyday commands

```bash
npm install -g @google/clasp        # v3+; commands renamed from v2 (open -> open-script)
clasp login                         # then enable the Apps Script API for your account:
                                    # https://script.google.com/home/usersettings
cd dashboard                        # one directory == one Apps Script project
clasp pull                          # ALWAYS pull and diff before you edit or push
clasp show-file-status              # dry run: exactly what push would upload
clasp push                          # replaces the remote project with local files
clasp open-script                   # open the IDE
clasp list-scripts                  # standalone projects only; bound scripts do not appear
```

There are no tests and no CI beyond GitGuardian secret scanning. `dashboard/Tests.js` has a
`runTests()` covering forecast math only, runnable from the IDE with no Drive or Xero access.

## Non-negotiables

- **`clasp push` replaces the remote project.** Files absent locally are deleted there. Never
  push without pulling and diffing first — on 2026-08-11 the deployed dashboard was ~1,030 lines
  ahead of the repo, including a whole permissions layer, and a blind push would have erased it.
  See SKILL.md §1, Deployment Invariant.
- **Never fetch code at runtime and execute it.** No `eval`, no `new Function` over a fetched
  string, no "loader" scripts. Code lives in the project and is deployed with clasp; secrets live
  in **Script Properties**, never in source. SKILL.md §1, No Remote Code Loading.
- **Verify AI review suggestions against the data before applying them.** A reviewer bot on PR #6
  asked for project abbreviations that matched neither the README nor the real Drive files;
  applying it would have decoupled generated budgets from live funding sources.
- **There is no `.gitattributes`.** clasp writes LF, Windows checks out CRLF, so whole files
  appear modified. Use `git diff --ignore-cr-at-eol` and never commit a line-endings-only diff.
- **Apps Script has one global scope per project.** Private helpers take a trailing underscore;
  only real entry points stay bare. SKILL.md §1, Global Scope Invariant.
- **Money changes need a human.** Anything altering published actuals, budgets, or what a funder
  is told is reviewed before it ships — say what will move and by how much.
- **Victor approves every commit and push.** Prepare the change, show the diffstat, then ask.

## Where things are

| | |
|---|---|
| Deep guide | [`.agents/skills/SKILL.md`](.agents/skills/SKILL.md) |
| Funding Cockpit dashboard | [`dashboard/`](dashboard/) — `README.md`, `GM_GUIDE.md`, `BUDGET_PROCEDURES_ADDENDUM.md` |
| Quarterly budget generation | [`project_reports/`](project_reports/) |
| Funding reports (Xero) | [`funding_reports/`](funding_reports/) — still contains a remote loader, see SKILL.md §5 |
| Chart-of-accounts helpers | `general_valid_accounts.js`, `variance_funding_source.js` (root) |
| Account-keyed budget/overhead logic | `create_xero_budget_project.js` (root) |
| Budgets (the actual data) | Google Drive `Budgets` folder — id `10105co6S5qHFSVVg0pb0fkoPidN3ScJZ` |

## Apps Script projects

| Project | Script ID | Kind |
|---|---|---|
| Funding Cockpit | `1L4ilqypO-LyLmY4Vyv6TwxjsUwH3d8x0cr54hIgkyU1bch7pQ4xAurVW` | standalone web app |
| quarterly_budgets | `18LAH4KWctMicL-lqNGTt8NkeXTzpiBVYOzwhYbUAgslxr3AEOZmQQdXy` | standalone |
| PROJECT_overview | not listed by clasp | bound to a spreadsheet |

Branches: work on a feature branch and open a PR against `dev`. `main` is the release branch.
