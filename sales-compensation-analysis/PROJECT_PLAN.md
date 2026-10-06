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
| 3 | The Assignment | ✅ | Four tasks: **3.1** June 2026 commissions for all 80 reps (AE, AM, frontline mgr, director, Head of Sales & AM), dynamic and traceable to source rows. **3.2** Data-quality exceptions ledger sized by records, $, reps, and treatment; includes payouts the plan allows but we would not sign off. Hint: one finding outweighs all others combined. **3.3** SMB AE accelerator 1.2x/1.3x → 1.3x/1.5x: paid vs repriced vs delta, with assumptions and surprises. **3.4** Slack scrape: escalate three items, write one response of ≤250 words. A1–A3 decided. |
| 4 | Deliverables | ✅ | **(1) Excel / Google Sheet model:** June 2026 commission for all 80 reps in USD **and local currency**, plus the exceptions ledger and the plan-change scenario. Inputs separate from calculations, built for the VP of Sales to open, no black-box or hardcoded logic, and the organization itself is evaluated. **(2) PDF summary, max 3 pages:** ledger findings ranked by $ impact; every assumption where the plan is silent, with a recommendation; the plan-change answer and surprises; the 3 Slack escalations and one reply. Open: Q10. |
| 5 | What the company is looking for | ✅ | Evaluation criteria: **Reconciliation discipline** (every exception named, sized, root-caused; stacked problems on one rep untangled). **Modeling craft** (consistent formulas, parameterised inputs; a teammate can rerun July or reprice without help). **Communication** (most important first; plain-language insight). **Comprehension** (strengths and weaknesses of the current plan; ability to own future plans). A5 decided. |
| 6 | Data files | ✅ | One workbook, **11 tabs** (known so far: tab 1 Sales Roster, tab 11 Slack Scrape). **All amounts are USD**: every facility bills in USD, so the revenue side never converts. Currency enters **only at payout**. The data deliberately contains more than needed, so we document which tabs we use and why. Appendix A explains what the tabs can't. Workbook arrives after Parts 7, 8 and the Appendix are reviewed. |
| 7 | Compensation plan | ✅ | Complete rule set = Section 07 + tabs 3 (FX), 4 (Ramp Terms), 5 (Compensation), 6 (Accelerators). Digest in `PLAN_RULES.md`. Answers Q1–Q5 and most of Q6/Q10. A6–A12 decided; A13 open until data; data watch list W1–W10; draft plan observations P1–P8. |
| 8 | Logistics | ✅ | Submit the model and the written summary **only** via the link in Ady's email; other formats or channels are not reviewed. → Final files: one `.xlsx` (no macros, cached values recalculated so it previews correctly) and one `.pdf` of ≤3 pages. Ady submits; I prepare the files. Q8 not answered → A14. |
| App. | Data dictionary | 🟨 | Defines the 11 tabs and how they join (see data inventory). Key rules: HR governs role, dates and manager line; the HR snapshot's Manager-ID held all month; the Head's Manager-ID has no row (expected); closed_by = AE, crm_owner = AM; segment doesn't decide the plan; the first worked shift is given, not derived; the shift extract covers Mar–Jun 2026 only; credits are USD and all approved types count. A13 and Q7 closed; A15–A16 proposed; W11–W18 added. |

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
| Q1 | Part 2 | Is Net Revenue calculated per shift, per hour, or per billed line? How are cancelled, no-show, partially worked, or adjusted/credited shifts treated? | Part 7, Appendix | Answered in part (7.1): NR = charge − pay on **worked** shifts, by date worked. How to identify cancelled, no-show and partial shifts → Appendix. |
| Q2 | Part 2 | What event marks an account as having "started producing" (first posted, claimed, worked, or billed shift)? | Part 7 | Answered (7.3, 7.4): two anchors. **Go-live** = effective_start_date of earliest Closed Won (drives Existing). **First worked shift** starts the AE windows. |
| Q3 | Part 2 | How long is the AE credit window, and how are partial months at its start and end handled? | Part 7 | Answered (7.4): 90 days = three 30-day windows from the first worked shift, not aligned to calendar months. |
| Q4 | Part 2 | When does the AM take over: at the start of production, or when the AE window ends? Can both be credited at once, or can neither be? | Part 7 | Answered (7.4, 7.7): no clean handoff. Once go-live is before the month, the AM is credited, while the AE may still be in their windows → **overlap by design** (see P1). |
| Q5 | Part 2 | Which date assigns revenue to June: shift date, invoice date, or payment date? How are late-arriving adjustments to earlier months handled? | Part 7, Part 6 | Answered (7.1, 7.2, 7.4): NR by date worked; GSV by date posted; credits by billing_month; F31-60D/F61-90D in the month the window closes. Nothing already paid is restated. |
| Q6 | Part 2 | How do roll-up quotas handle mid-month hires, departures, transfers, and open territories? | Part 7, Part 6 | Answered in part (7.8, 7.11): whole Manager chain; quota from tab 5; leavers' revenue still rolls up. Mid-month transfers → A13. |
| Q7 | Part 2 | Do we have the June statements already sent (and the old spreadsheet), so we can reconcile our numbers against what was paid? | Part 6 | Answered (Appendix): no tab holds the June statements or the old spreadsheet. The only evidence of what was paid is in the Slack scrape. A2/A3 sizing is therefore vs a naive calculation only. |
| Q8 | Part 2 | When our June numbers differ from what was paid, what is the policy for correcting it (true-up next cycle, off-cycle payment, recovering an overpayment)? | Part 7, Part 8 | Not answered by the plan or logistics → moved to A14. |
| Q9 | Part 3 | What tool must the dynamic model use (Excel, Google Sheets, SQL, Python)? Is a single workbook expected? | Part 4, Part 8 | Answered (Part 4): a single Excel/Google Sheet model with live formulas. See A4. |
| Q10 | Part 4 | Which FX rate converts USD commission to local currency (June average, month-end, or payment date), and where does it come from? Do we round the USD commission first and then convert, or convert unrounded and round in local currency? *Part 6 narrowed this: conversion happens once, at payout only.* | Part 7, Appendix | Answered in part (7.13): tab 3 rate, local units per USD, multiply. Rounding order → A12. |

## Ambiguity log

| ID | Source (part / file) | Ambiguity | Options & $ impact | Recommendation | Decision | Status |
|----|----------------------|-----------|--------------------|----------------|----------|--------|
| A1 | Part 3 (3.1 vs 3.2) | 3.1 says to apply the plan "exactly as written", but 3.2 asks us to flag what the plan would pay that we would not sign off. Which number is the June payout? | (a) Plan as written only. (b) Recommended number only. (c) Both, side by side. $ impact to be sized. | **(c)** Data errors get corrected in the calculation, because they are not the plan. Policy concerns are kept as a separate, visible hold/adjustment layer, so the plan-as-written figure and the sign-off figure can both be traced. | ✅ Approved as recommended (2026-10-05) | Decided |
| A2 | Part 3 (3.2) | "Dollars moved" needs a baseline. | (a) Versus a naive calculation on the raw data. (b) Versus the June statements already paid. | **(a)** as the primary sizing, because it is reproducible. Add (b) where the paid statements exist. | ✅ Approved as recommended (2026-10-05) | Decided |
| A3 | Part 3 (3.3) | "June as actually paid" could mean the statements that went out, or our corrected June. | (a) Statements as issued. (b) Our corrected calculation. | **(b)** as the baseline for repricing, because the issued statements may contain errors that would distort the delta. Show (a) alongside if the data exists. | ✅ Approved as recommended (2026-10-05) | Decided |
| A4 | Part 4 | Excel or Google Sheet? | (a) Excel .xlsx with live formulas, which also opens in Google Sheets. (b) A native Google Sheet. | **(a)**: I can build and verify it here and it works in both tools; you can import it into Google Drive if you prefer. | ✅ Approved: Excel (2026-10-06) | Decided |
| A5 | Part 5 | When two or more problems stack on one rep, how is each one's $ impact sized? Sizing them one at a time can double-count or miss interaction effects (e.g. two fixes that each push a rep across an accelerator tier). | (a) Standalone: each fix alone vs naive. (b) Sequential waterfall in a fixed order. (c) Both. | **(c)**: the ledger uses a sequential waterfall in a stated order, so the steps add up exactly to the final change; the per-rep view also shows each fix's standalone impact and the interaction difference. | ✅ Approved as recommended (2026-10-06) | Decided |
| A6 | 7.2 | A June approved credit on an account that is **not Existing in June** (e.g. went live in June) has no Existing revenue to reduce. | (a) Apply it to crm_owner anyway. (b) No effect on anyone; ledger it. | **(b)**: 7.2 says credits reduce Existing only and new-logo is measured before credits. Ledger it so finance can see the credit wasn't charged to anyone. | ✅ Approved as recommended (2026-10-06) | Decided |
| A7 | 7.2, 7.9 | Credits can push an AM's June Existing revenue below zero → **negative commission**. The plan says "no floor" (about quota), but is silent on clawback. | (a) Pay a negative amount (claw back). (b) Floor the month at $0, no carry-forward. (c) Floor at $0 and carry the deficit forward. | **(b)** for June, flagged for leadership: clawing back pay through a statement needs a written policy (7.14). The roll-up still uses the true (negative) figure, because managers carry what reps are paid on → raise if this happens. | ✅ Approved as recommended (2026-10-06) | Decided |
| A8 | 7.6 | An Applied Science tag appears **after** windows were already paid on NR. Is the account GTM for the whole 90 days, or only from the tagged shift on? | (a) GTM for all windows; earlier months not restated. (b) GTM only from the first tagged shift. | **(a)**: GTM is an account property ("flips an account"), so all windows are compared on the same basis (GSV vs GSV). Paid months are not restated (7.4). Ledger the comparison. | ✅ Approved as recommended (2026-10-06) | Decided |
| A9 | 7.6 | GTM GSV **posted in a month that already closed** (e.g. posted in May, first worked shift in June) falls into F1-30D, but its recognition month (May) has been paid. | (a) Catch it up on the June statement. (b) Lose it (May is closed). | **(a)**: the credit only became knowable once a shift was worked, so paying it in June is first recognition, not a restatement. Ledger it with $. | ✅ Approved as recommended (2026-10-06) | Decided |
| A10 | 7.9, 7.10, 7.12 | A rep whose ramp **ends partway through June**: are they "ramping" for June (no accelerators, no SPIFF)? | (a) Ramping if any June day is in a ramp month. (b) Ramping only if the floor actually binds. (c) Split the month by day. | **(a)**: the plan treats ramp as a status; a one-day test is objective and repeatable. Show the $ difference vs (b) for affected reps. | ✅ Approved as recommended (2026-10-06) | Decided |
| A11 | 7.11 | A rep who **changed role in June**: which role's rate, bands, quota and rollup line apply? | (a) Roster role at month end, for the whole month. (b) Split credit by date across both roles. | **(a)**: 7.11 says quota does not prorate for a role change, which implies one role per month. Flag each case and show the $ under (b). | ✅ Approved as recommended (2026-10-06) | Decided |
| A12 | 7.13 | Rounding: when do we round to cents? | (a) Full precision throughout; round USD to cents per rep; convert; round local to the currency's minor unit. (b) Round at every step. | **(a)**: rounding once avoids pennies piling up across bands and windows; state it in the README. | ✅ Approved as recommended (2026-10-06) | Decided |
| A13 | 7.7, 7.8, A-2, A-8 | `crm_owner_id` and the Manager chain are snapshots from the export. | (a) Use the export as is. (b) Rebuild June ownership from history. | **Resolved by the Appendix:** A-2 says the Manager-ID shown held for the whole month and there is no manager history; A-8 gives only the current owner, with no ownership history to rebuild from. Use (a). Ledger any contrary evidence (e.g. Slack) without changing the calculation. | Resolved by Appendix | Closed |
| A14 | Part 2, 7.14 | June statements have already been paid. When our recalculation differs, how is the difference settled? The plan is silent. | (a) Underpayments: pay on the July statement or off-cycle. (b) Overpayments: deduct automatically from July. (c) Overpayments: hold for a written leadership decision. | **(a) + (c)**: pay underpayments promptly (off-cycle if material). Overpayments go to sales leadership for a written decision (recover, spread, or forgive), never deducted automatically, consistent with 7.14. State this in the PDF. | ✅ Approved as recommended (2026-10-06) | Decided |
| A15 | A-9, 7.4 | `account_first_worked_shift_date` is given and we must not infer it. What if it conflicts with the data (different values on rows of the same account, or a worked shift in the extract dated before it)? | (a) Use the field as given; if an account has several values, use the earliest. (b) Override it with the observed first worked shift. | **(a)**: the dictionary explicitly says not to infer it. Each conflict goes in the ledger with the $ it would move under (b). | Pending | Proposed |
| A16 | A-1, A-2 | The Sales Roster (tab 1) and the HR Export (tab 2) may disagree on role, manager, dates or status. | (a) HR governs role, dates and Manager-ID; the roster defines the 80 in scope and the CRM id. (b) The roster governs. | **(a)**: A-2 says Current Role 'decides which plan the rep is paid on', and the roster 'is not guaranteed to agree with HR'. Ledger every disagreement with its $ effect. | Pending | Proposed |

## Proposed model structure (Part 4)

| Layer | Tabs | Rule |
|-------|------|------|
| Guide | `README`: purpose, how to trace a payout, tab map, colour legend | First tab the VP sees |
| Inputs | Raw data tabs exactly as received; `Plan_Parameters` (rates, tiers, accelerators, windows, caps); `FX_Rates`; `Treatments` (exception decisions as flags) | Never edited by formulas; parameters live in exactly one place |
| Calculations | Shift-level crediting → rep-month attainment → AE / AM / manager / director / Head roll-ups | Formulas only, referencing inputs; every row carries source row IDs |
| Outputs | `Summary` (80 reps, USD + local, as-written vs recommended); `Exceptions_Ledger`; `Scenario_Accelerator` | Scenario driven by a parameter switch, not a copied model |
| Checks | `Reconciliation`: totals tie to source, row counts, every rep present once | All checks must show PASS before sign-off |

## Design implications from Part 5

- **Ledger columns:** ID · finding · root cause · records touched (source row IDs) · reps affected · $ naive → $ treated → $ moved · treatment and why · category (data error / plan-as-written but not signed off) · decision owner.
- **Stacked issues:** a per-rep waterfall (naive → fix 1 → fix 2 → … → final), using the A5 sizing method.
- **Rerun July:** the comp month is a single input; credit windows, FX and quotas are all looked up from it, with no dates inside formulas. The README includes a step-by-step "run next month" runbook.
- **Communication:** every output tab and the PDF start with the headline number and the top finding, then the detail.
- **Comprehension:** a `Plan_Observations` tab (strengths, weaknesses, incentive risks such as behaviours the plan rewards, cliffs, gaps or overlaps at handoffs, roll-up effects) with recommended plan changes. The best points go into the PDF.

## Data watch list (from Part 7 — check when the workbook arrives)

| ID | What to check | Why it matters |
|----|---------------|----------------|
| W1 | Shift status values: posted, claimed, worked, cancelled, no-show, partial | Only worked shifts carry NR; every posted shift carries GSV |
| W2 | Opportunity rows with `closed_won_date` but stage ≠ Closed Won | Reversed deals must not set go-live or create credit (7.3) |
| W3 | Account properties (rates, terms, segment, type) that differ across opportunities | 7.3 says they're identical; differences are a data finding |
| W4 | Blank `closed_by_owner_id`; closers who left or aren't AEs | Nobody credited / manager carries / rate to apply |
| W5 | Net terms values outside 15/30/45/60 | No multiplier defined |
| W6 | Applied Science tags, especially on accounts with heavy **unfilled** postings | GSV pays on shifts never worked → gaming risk; strong candidate for the "biggest finding" |
| W7 | Credits: denied vs approved, billing_month outside June, credits on non-Existing accounts | A6, A7 |
| W8 | Roster: hire/leave dates, Current Role Start Date, rehires, mid-month role changes, Manager chain loops or gaps | Quota proration, ramp, roll-ups, A10, A11 |
| W9 | Shifts worked **before** go-live or before Closed Won | They start AE windows, but the account isn't live yet |
| W10 | Duplicate shift or credit rows; charge rate below pay rate (negative margin) | Classic double-pay and sign errors |

## Draft plan observations (feed `Plan_Observations` and the PDF)

| ID | Observation | Effect |
|----|-------------|--------|
| P1 | AE windows (from first shift) and AM Existing credit (from go-live) **overlap** | The same dollars can pay twice in an account's early months |
| P2 | Ramp floor rolls up to managers | Managers are paid on revenue nobody produced |
| P3 | GTM pays on GSV of **posted** shifts, filled or not | Rewards posting volume, not revenue; easy to inflate |
| P4 | Improvement test on windows 2–3 | Rewards a soft first window (sandbagging) to show improvement later |
| P5 | Net 15 at 120% | Pays a premium on revenue that isn't higher; worth checking how many accounts claim Net 15 |
| P6 | Retention SPIFF is all-or-nothing at 100%, and the AM controls credit timing | A small credit moved between months can flip the SPIFF |
| P7 | Window credits wait until the window closes | Pay lags performance by up to 2 months; harder to explain to reps |
| P8 | Closer keeps credit after leaving ownership | Clear and fair to hunters, but the AE and AM can be credited simultaneously (P1) |

## Data inventory (from Appendix A — confirm when the workbook arrives)

| Tab | Rows | Contents | Used for | Watch |
|-----|------|----------|----------|-------|
| 1. Sales Roster | 80 | Reps in scope; unique names; `crm_owner_id` | Population for 3.1; links reps to CRM | Disagreements with HR (A16) |
| 2. HR Export | 1,080 | Role, employment dates (inclusive), role dates, Manager-ID, Country Code | Plan role, quota proration, ramp clock, roll-up chain, payout currency | Join key to roster (W11); chain gaps/loops |
| 3. Exchange Rates | 196 | Local units per 1 USD | USD → local payout | Country → currency mapping (W12) |
| 4. Ramp Terms | 17 roles | Start Date Ramp (cols B–D); New Report Ramp (E–G) not used by any role | Ramp floor | Months in B–D vs plan's 2 defined months (W13) |
| 5. Compensation | 17 roles | Headcount, # of Reports, quota, rates; Metric II for AMs and the Head | Rates, quotas, manager quotas | Headcount vs roster/HR (W14) |
| 6. Accelerators | — | AE bands and multipliers; AM retention SPIFF | Bands, SPIFF; 3.3 scenario | — |
| 7. Opportunity to ID Mapping | 761 | CRM ↔ app ↔ billing ids | Joins shifts and credits to accounts | 761 vs 815 opps; duplicates and gaps (W15) |
| 8. Opportunity Export | 815 | Stage, closed_won_date, effective_start_date, closed_by (AE), crm_owner (AM), rates, terms, segment | Go-live, AE/AM attribution, net terms | W2–W5; segment is informational only |
| 9. App Export | 107,052 | Posted shifts; `worked_shift`; `shift_worked_date`; `account_first_worked_shift_date`; Applied Science tag | NR, GSV, windows, GTM | Extract window Mar–Jun (W16); flag/date consistency (W17); A15 |
| 10. Billing Export | 1,089 | Credit tickets (USD), status, billing_month, ticket_type | Existing NR deductions; SPIFF May/June | Statuses other than approved/denied (W18) |
| 11. Slack Scrape | — | Messages from June to mid-July | 3.4 | Only evidence of what was actually paid |

**Join path:** shift (app id) → tab 7 → account / opportunities (tab 8) → `closed_by_owner_id` / `crm_owner_id` → tab 1 roster → tab 2 HR (role, manager, country) → tab 3 FX. Credits follow the same path via billing id.

### Watch list additions from Appendix A

| ID | What to check | Why it matters |
|----|---------------|----------------|
| W11 | The key that joins the roster to HR (an id, or name only?) | A bad join silently mis-assigns role, manager or currency |
| W12 | How Country Code maps to a tab 3 rate (shared currencies, missing countries) | Wrong local payout |
| W13 | Whether tab 4 has a 3rd ramp month although 7.10 defines only months 1–2 | Possible plan/tab conflict |
| W14 | Tab 5 headcount and # of Reports vs the actual roster and HR chain | Manager quota is taken from tab 5, not from actual team size |
| W15 | Tab 7: ids mapped to more than one account, unmapped ids, many opportunities per account | **Duplicate mapping multiplies revenue**; strong "biggest finding" candidate alongside W6 |
| W16 | Unworked posted shifts outside Mar–Jun, and GTM postings before 1 March | GSV for F1-30D may be incomplete; data limitation to ledger |
| W17 | `worked_shift` TRUE without a date, FALSE with a date; first-shift date inconsistent within an account | Revenue on or off by mistake; A15 |
| W18 | Credit statuses other than approved/denied (e.g. pending); billing_month outside May–June | Only approved credits count; May credits matter for the SPIFF |

## Decision log

| Date | Decision | Rationale |
|------|----------|-----------|
| 2026-10-06 | A14 approved; A13 closed by Appendix (no ownership or manager history exists) | See ambiguity log |
| 2026-10-06 | A6–A12 approved as recommended | See ambiguity log |
| 2026-10-06 | A5 approved: sequential waterfall + standalone impact | See ambiguity log |
| 2026-10-06 | Deliverable built as Excel .xlsx with live formulas (A4) | Can be built and verified here; also opens in Google Sheets |
| 2026-10-05 | A1–A3 approved as recommended | See ambiguity log |
| 2026-10-05 | Project lives in `sales-compensation-analysis/` within this repo | Synthetic data, safe to store; keeps portfolio README untouched |
