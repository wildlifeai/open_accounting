# Budget procedures - dashboard addendum

The Funding Cockpit dashboard reads the same budget Gsheets and Xero data your existing
procedures already produce, so **most of the Create / Maintain / Archive workflow is
unchanged**. There are only a few small additions, listed here so they can be folded into
the three Google Docs.

The dashboard depends on conventions you already follow:
- the funding source Gsheet is **named exactly as the Xero "Funding source" tracking value**
  (e.g. `WW_25_TOI`);
- it lives in a `secured/` or `proposed/` folder under its project;
- the `Budget` tab has `Start`, `End`, `Cost`, `Income`, `Milestone`, `Xero Inventory Item`;
- Xero transactions are tagged with both tracking categories (*Projects* and
  *Funding source*) and the product/service item.

If those hold, a budget appears on the dashboard automatically. The additions below cover the
one new field and a few clarifications.

---

## Create a Funding Source Budget - additions

**Phase 1, step 4 (itemise the Budget tab) - add:**
> If a budget line relates to a **different project** than the folder this budget sits under,
> fill the optional **`Project`** column on that line with the project's exact Xero *Projects*
> tracking value (e.g. `General`). Leave it blank for lines that belong to this budget's own
> project - blank means "use the parent project folder". This is what lets one funding source
> (e.g. `WW_25_TOI` under Wildlife Watcher) book its general-management lines to General while
> the rest stay with Wildlife Watcher.

**Phase 3 (Xero tracking) - add a reminder:**
> For the dashboard to split actuals correctly, every transaction must carry **both** tracking
> categories - *Projects* and *Funding source* - as well as the product/service item. A
> transaction missing the *Projects* tag won't be attributed to a project on the dashboard.

No other Create changes. The `*Account` column remains optional; include it if you want
account-level detail, omit it otherwise - the dashboard handles both.

---

## Maintain a Funding Source Budget - **significant change**

The dashboard takes over the two jobs the `Budget, Actual, Forecast Tracking` tab used to do
by hand, so this procedure simplifies substantially. **The per-funding-source tracking tab
and its "Xero Sync" are retired.** Replace the Maintain procedure body with the following.

**Phase 1: Reconcile actuals (now in Xero, not in the sheet)**
> Actuals are read **live from Xero** by the dashboard - you no longer sync transactions into
> the funding-source Gsheet. To reconcile: open the dashboard's **Quarterly tracking** tab,
> select the funding source, and compare the *Actual to date* totals with Xero. If something
> looks wrong it's a coding issue - find the transaction in Xero and check it carries the
> right *Projects* tag, *Funding source* tag, and product/service item. The dashboard
> refreshes every ~6 hours, or immediately with **Refresh now**.

**Phase 2: Update the forecast (in the funding source's own `Forecast` tab)**
> SUPERSEDED as of 2026-08-11. Forecasts are **not** held centrally. The central sheet is
> *Cockpit Settings* (Permissions only); each funding source carries its own `Forecast` tab,
> read by `BudgetReader.parseForecastTab_` on every refresh. See
> [BUDGET_SHEET_TEMPLATE.md](BUDGET_SHEET_TEMPLATE.md) for its exact layout.
>
> There is no quarterly data-entry obligation: a quarter with no `Forecast` entry falls back to
> the budget baseline. Record an override only when you know something the budget does not — a
> delayed hire, a grant ending early, a re-profiled milestone — and use the `Comments` column to
> say why. The dashboard's **Quarterly tracking** tab shows actuals for past quarters, the
> override where one exists, and the baseline everywhere else.

**What stays the same**
> The `Budget` tab remains the **frozen baseline** (the approved plan). Only change it for a
> genuine re-budget, not for routine forecasting - the dashboard compares your forecast and
> actuals against it to show variance.

**Optional cleanup**
> Existing `Budget, Actual, Forecast Tracking` tabs can be left in place for history or
> removed; the dashboard ignores them. The `Submitted_budget` tab is unaffected.

---

## Archive a Funding Source Budget - additions

**No procedural change needed - add a confirming note:**
> Following this procedure removes the funding source from the dashboard automatically: the
> dashboard only reads the `secured/` and `proposed/` folders and ignores anything named with
> the `Z_ARCH_` prefix. Once the Gsheet is renamed `Z_ARCH_...` and moved to the archived
> folder, it stops appearing at the next refresh. Keep the product/service item codes
> unchanged (as the procedure already states) so historical actuals still reconcile.
