# June 2026 Compensation Plan — Rules Digest

A plain-language restatement of Section 07, with the plan tabs it points to
(4. Ramp Terms, 5. Compensation, 6. Accelerators, 3. Exchange Rates).
Section references in brackets. Where the plan is silent, see the ambiguity log in
`PROJECT_PLAN.md`; nothing here fills a gap silently.

## 1. Revenue measures

| Measure | Definition | Recognised on | Used for |
|---------|-----------|---------------|----------|
| **Net Revenue (NR)** | charge rate − pay rate, **worked shifts only** | date **worked** | AE (non-GTM) and AM credit [7.1] |
| **Gross Services Value (GSV)** | charge rate on **every posted shift**, filled or not | date **posted** | AE credit on GTM accounts only [7.1, 7.6] |
| **Billing credits** | approved tickets only; denied tickets change nothing | `billing_month` | Reduce **AM (Existing) only**; never AE [7.2] |

## 2. What an account is [7.3]

- **Go-live** = `effective_start_date` of the account's **earliest Closed Won** opportunity.
  Later expansions or renewals do not reset it.
- `opportunity_stage` decides whether a deal is won. A `closed_won_date` on a row that is not
  Closed Won is left over from a reversed deal, **not a win**.
- Rates, net terms, segment and account type belong to the **account** and should match
  across all its opportunities. Any mismatch is a data-quality finding.

> Two different anchors: **go-live** (contract date) decides when an account becomes
> *Existing* for the AM; the **first worked shift** starts the AE's 90-day windows.

## 3. Account Executive — New Customer credit [7.4, 7.5, 7.6]

**Who:** `closed_by_owner_id` as in the export, even if they no longer own the account.
If it's blank, **nobody** is credited (not even a manager).

**Windows** run in days from the account's first worked shift (day 1 = that date):

| Window | Days | What is credited | When it is credited | Rate |
|--------|------|------------------|---------------------|------|
| F1-30D | 1–30 | every dollar | month each shift was worked | full |
| F31-60D | 31–60 | max(0, this window − F1-30D) | in one piece, in the month the window **closes** (day 60) | half |
| F61-90D | 61–90 | max(0, this window − best of the two earlier windows) | month it closes (day 90) | half |

- The half is applied **to the credit**, so quota attainment and accelerator bands see the halved number.
- Nothing already paid is restated. A window that goes backwards credits $0.
- Windows that closed or ran in earlier months still need **April–May shift history** for the
  comparisons.

**Net terms** scale the revenue (not the rate), on new-logo NR only:
Net 15 → 120% · Net 30 → 100% · Net 45 → 90% · Net 60 → 70%.
Never applied to Existing revenue or to GTM.

**GTM accounts** (any shift tagged *Applied Science* flips the account):
- Same windows and improvement test, but measured on **GSV by posted date**.
- Windows only have an end date: shifts **posted before** the first worked shift fall into F1-30D.
- Pays 3% instead of the role rate. Converted into credited revenue as
  `GSV × 3% ÷ role rate`, so it retires quota and fills bands like any other credit.
- AE only. AMs on GTM accounts are paid on ordinary NR.

## 4. Account Manager — Existing Customer credit [7.7, 7.2]

- An account is **Existing in June** if its go-live is **before 1 June 2026**.
- Credit goes to `crm_owner_id`.
- June figure = NR of shifts worked in June on that book − approved credits with `billing_month` = June.
- No net-terms adjustment, no GSV, **no accelerators**.

> As written, an account that went live in May can be credited to the **AE (windows)** and the
> **AM (Existing)** in the same month. That overlap is by design, not a data error.

## 5. Managers, directors, Head of Sales & AM [7.8]

- Carry the **sum of what everyone beneath them in the Manager chain is paid on**
  (after any ramp floor), not just direct reports.
  - AE managers and the Director of Sales → New.
  - AM managers and the two AM Directors → Existing.
  - Head of Sales & AM → both, each at its own rate (tab 5).
- Quota comes from tab 5 (built from reports' quotas, not their production).
- No accounts, no accelerators, no ramp.
- If a rep can't be paid (e.g. they left), **the manager still carries the revenue**.
  Blank `closed_by_owner_id` → nobody carries it.

## 6. Turning credit into commission [7.9]

- Commission = credited revenue × role's Standard Payout rate, with accelerator bands (tab 6)
  applied to the **incremental dollars in each band**.
  - Example, SMB AE at $45k: $30k at standard + $10k at band I multiplier + $5k at band II.
- Bands are **absolute dollars**, never prorated (even for a partial month).
- No floor, no cap. Quota only measures attainment.
- Accelerators: **SMB AE and ENT AE only**, and **not while ramping**.

## 7. Ramp [7.10]

- Roles that ramp: SMB AE, ENT AE, SMB AM, ENT AM.
- Clock starts at **Current Role Start Date** (restarted by a promotion into a ramped role or a rehire).
  Month 1 = days 1–30, month 2 = days 31–60 (percentages on tab 4).
- Ramp credit is a **floor**, not a bonus: credited = max(produced, ramp credit).
- Ramp credit is built day by day: each June day in a ramp month earns
  `ramp % for that month × monthly quota ÷ 30` (June has 30 days). Assumes Net 30, so no terms adjustment.

## 8. Joiners, leavers, role changes [7.11]

- Quota prorates by **calendar days employed** in June. It does **not** prorate for a role change.
- A rep is credited only for credit recognised **on a date they were employed**
  (shift date for F1-30D and Existing; window-close date for F31-60D/F61-90D).
  Credit they can't receive still rolls up to their manager.

## 9. Retention SPIFF [7.12]

- **AM individual contributors only**; amounts on tab 6. Not payable while ramping.
- Book = accounts the AM holds in June that went live **before 1 May**.
- Pays if June ECNR on that book ≥ 100% of May ECNR on the same book, both net of approved
  credits for their own month. May = $0 → no SPIFF.

## 10. Currency [7.13]

- Everything is computed in USD. Local payout = USD commission × tab 3 rate
  (rate = local units per 1 USD).

## 11. What governs [7.14]

- Only Section 07, the plan tabs, and adjustments **approved in writing by sales leadership**.
- Slack messages, promises and outside conventions do not change pay. We can still flag
  payouts we would not sign off.
