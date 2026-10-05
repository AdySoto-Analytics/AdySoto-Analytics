# Sales Compensation Analysis — Project Plan & Tracker

A living document to trace the project together. Each requirement part is reviewed and
agreed **before** any calculation begins.

## Goals

1. **Accuracy** — produce a commission number that is right to the dollar.
2. **Design** — understand what makes an incentive system effective.
3. **People** — navigate real-life, high-stakes situations with the individuals whose
   compensation we calculate.

## Working rules

- No silent assumptions: every ambiguity is logged with options, dollar impact, and a
  recommended treatment. Ady decides; affected numbers stay **provisional** until then.
- Every payout is traceable: inputs → plan rule → rate/tier → result.
- Totals are reconciled to source data; rounding rules are stated explicitly.
- Data in this folder is synthetic test data built to stress the tracker.

## Phase 1 — Review the requirements

Status key: ⬜ Not started · 🟨 In discussion · ✅ Agreed

| Part | Topic | Status | Notes / key takeaways |
|------|-------|--------|-----------------------|
| 1 | Goals | ✅ | Three goals above; ambiguity-handling rule agreed. |
| 2 | Background | ✅ | Healthcare staffing marketplace; Net Revenue = charge rate − pay rate on worked shifts. AEs paid on New Customer NR (limited window after account starts producing); AMs on Existing Customer NR afterwards; managers/directors on roll-up quotas. Comp month **June 2026**; statements already sent and questioned; rebuilding from scratch. Open questions Q1–Q8. |
| 3 | The Assignment | 🟨 | Four tasks: **3.1** June 2026 commissions for all 80 reps (AE, AM, frontline mgr, director, Head of Sales & AM), dynamic and traceable to source rows. **3.2** Data-quality exceptions ledger sized by records, $, reps, and treatment; includes payouts the plan allows but we would not sign off. Hint: one finding outweighs all others combined. **3.3** SMB AE accelerator 1.2x/1.3x → 1.3x/1.5x: paid vs repriced vs delta, with assumptions and surprises. **3.4** Slack scrape: escalate three items, write one response of ≤250 words. Open: A1–A3, Q9. |
| 4 | Deliverables | 🟨 | **(1) Excel / Google Sheet model:** June 2026 commission for all 80 reps in USD **and local currency**, plus the exceptions ledger and the plan-change scenario. Inputs separate from calculations, built for the VP of Sales to open, no black-box or hardcoded logic, and the organization itself is evaluated. **(2) PDF summary, max 3 pages:** ledger findings ranked by $ impact; every assumption where the plan is silent, with a recommendation; the plan-change answer and surprises; the 3 Slack escalations and one reply. Open: Q10, A4. |
| 5 | What the company is looking for | ⬜ | |
| 6 | Data files | ⬜ | |
| 7 | Compensation plan | ⬜ | |
| 8 | Logistics | ⬜ | |
| App. | Data dictionary | ⬜ | |

## Later phases (refined once Phase 1 is agreed)

| Phase | Description | Status |
|-------|-------------|--------|
| 2 | Data profiling & quality checks (feeds 3.2 ledger) | ⬜ |
| 3 | Commission calculation, all 80 reps (3.1) | ⬜ |
| 4 | Exceptions ledger, sized and treated (3.2) | ⬜ |
| 5 | Plan-change model: SMB AE accelerator (3.3) | ⬜ |
| 6 | Slack triage: 3 escalations + one response of ≤250 words (3.4) | ⬜ |
| 7 | Deliverables: Excel model + PDF summary of 3 pages max (Part 4) | ⬜ |

## Open questions (to verify against later parts)

Questions raised while reading the requirements. If a later part answers one, we close it.
If it stays unclear, it moves to the ambiguity log with a recommended treatment.

| ID | Raised in | Question | Expected answer in | Status |
|----|-----------|----------|--------------------|--------|
| Q1 | Part 2 | Is Net Revenue calculated per shift, per hour, or per billed line? How are cancelled, no-show, partially worked, or adjusted/credited shifts treated? | Part 7, Appendix | Open |
| Q2 | Part 2 | What event marks an account as having "started producing" (first posted, claimed, worked, or billed shift)? | Part 7 | Open |
| Q3 | Part 2 | How long is the AE credit window, and how are partial months at its start and end handled? | Part 7 | Open |
| Q4 | Part 2 | When does the AM take over: at the start of production, or when the AE window ends? Can both be credited at once, or can neither be? | Part 7 | Open |
| Q5 | Part 2 | Which date assigns revenue to June: shift date, invoice date, or payment date? How are late-arriving adjustments to earlier months handled? | Part 7, Part 6 | Open |
| Q6 | Part 2 | How do roll-up quotas handle mid-month hires, departures, transfers, and open territories? | Part 7, Part 6 | Open |
| Q7 | Part 2 | Do we have the June statements already sent (and the old spreadsheet), so we can reconcile our numbers against what was paid? | Part 6 | Open |
| Q8 | Part 2 | When our June numbers differ from what was paid, what is the policy for correcting it (true-up next cycle, off-cycle payment, recovering an overpayment)? | Part 7, Part 8 | Open |
| Q9 | Part 3 | What tool must the dynamic model use (Excel, Google Sheets, SQL, Python)? Is a single workbook expected? | Part 4, Part 8 | Answered (Part 4): a single Excel/Google Sheet model with live formulas. See A4. |
| Q10 | Part 4 | Which FX rate converts USD to local currency (June average, month-end, or payment date)? Where does it come from, and do we round before or after converting? | Part 6, Part 7, Appendix | Open |

## Ambiguity log

| ID | Source (part / file) | Ambiguity | Options & $ impact | Recommendation | Decision | Status |
|----|----------------------|-----------|--------------------|----------------|----------|--------|
| A1 | Part 3 (3.1 vs 3.2) | 3.1 says to apply the plan "exactly as written", but 3.2 asks us to flag what the plan would pay that we would not sign off. Which number is the June payout? | (a) Plan as written only. (b) Recommended number only. (c) Both, side by side. $ impact to be sized. | **(c)** Data errors get corrected in the calculation, because they are not the plan. Policy concerns are kept as a separate, visible hold/adjustment layer, so the plan-as-written figure and the sign-off figure can both be traced. | ✅ Approved as recommended (2026-10-05) | Decided |
| A2 | Part 3 (3.2) | "Dollars moved" needs a baseline. | (a) Versus a naive calculation on the raw data. (b) Versus the June statements already paid. | **(a)** as the primary sizing, because it is reproducible. Add (b) where the paid statements exist. | ✅ Approved as recommended (2026-10-05) | Decided |
| A3 | Part 3 (3.3) | "June as actually paid" could mean the statements that went out, or our corrected June. | (a) Statements as issued. (b) Our corrected calculation. | **(b)** as the baseline for repricing, because the issued statements may contain errors that would distort the delta. Show (a) alongside if the data exists. | ✅ Approved as recommended (2026-10-05) | Decided |
| A4 | Part 4 | Excel or Google Sheet? | (a) Excel .xlsx with live formulas, which also opens in Google Sheets. (b) A native Google Sheet. | **(a)**: I can build and verify it here and it works in both tools; you can import it into Google Drive if you prefer. | Pending | Proposed |

## Proposed model structure (Part 4)

| Layer | Tabs | Rule |
|-------|------|------|
| Guide | `README`: purpose, how to trace a payout, tab map, colour legend | First tab the VP sees |
| Inputs | Raw data tabs exactly as received; `Plan_Parameters` (rates, tiers, accelerators, windows, caps); `FX_Rates`; `Treatments` (exception decisions as flags) | Never edited by formulas; parameters live in exactly one place |
| Calculations | Shift-level crediting → rep-month attainment → AE / AM / manager / director / Head roll-ups | Formulas only, referencing inputs; every row carries source row IDs |
| Outputs | `Summary` (80 reps, USD + local, as-written vs recommended); `Exceptions_Ledger`; `Scenario_Accelerator` | Scenario driven by a parameter switch, not a copied model |
| Checks | `Reconciliation`: totals tie to source, row counts, every rep present once | All checks must show PASS before sign-off |

## Decision log

| Date | Decision | Rationale |
|------|----------|-----------|
| 2026-10-05 | A1–A3 approved as recommended | See ambiguity log |
| 2026-10-05 | Project lives in `sales-compensation-analysis/` within this repo | Synthetic data, safe to store; keeps portfolio README untouched |
