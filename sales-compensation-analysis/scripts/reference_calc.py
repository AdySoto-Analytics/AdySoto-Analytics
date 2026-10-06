"""Reference calculation of June 2026 commissions (independent check for the Excel model).

Implements Section 07 as summarised in PLAN_RULES.md. Options for open ambiguities are
exposed as keyword arguments so their dollar impact can be sized.
"""
import sys
from pathlib import Path
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from load import load

MONTH_START = pd.Timestamp("2026-06-01")
MONTH_END = pd.Timestamp("2026-06-30")
DAYS_IN_MONTH = 30
TERMS_MULT = {15: 1.2, 30: 1.0, 45: 0.9, 60: 0.7}
GTM_RATE = 0.03
HALF = 0.5


def build(d=None, *, gtm_exclude=(), accel_smb=(1.2, 1.3), go_live="earliest_cw",
          stage_rule="stage", role_start_override=None, ramp_status="any_day",
          catch_up_gtm=True, negative_am="floor0"):
    d = d or load()
    r, h, fx = d["1. Sales Roster"], d["2. HR Export"], d["3. Exchange Rates"]
    comp, acc = d["5. Compensation"], d["6. Accelerators"]
    m, o, a, b = (d["7. Opportunity to ID Mapping"], d["8. Opportunity Export"],
                  d["9. App Export"], d["10. Billing Export"])

    # ---------- people ----------
    p = r.merge(h.drop(columns="src_row"), on="hr_id")
    p["role"] = p["Current Role"]
    p["role_start"] = p["Current Role Start Date"]
    if role_start_override:
        for nm, dt in role_start_override.items():
            p.loc[p.name == nm, "role_start"] = pd.Timestamp(dt)
    p["emp_start"] = p["Start Date"]
    p["emp_end"] = p["End Date"]
    cr = comp.set_index("Role")
    p["quota"] = p.role.map(cr["Monthly Quota Target"]).astype(float)
    p["rate"] = p.role.map(cr["Standard Payout"]).astype(float)
    days = pd.date_range(MONTH_START, MONTH_END)

    def employed(row, dt):
        return (row.emp_start <= dt) and (pd.isna(row.emp_end) or dt <= row.emp_end)

    p["days_employed"] = p.apply(lambda row: sum(employed(row, dt) for dt in days), axis=1)
    p["quota_prorated"] = p.quota * p.days_employed / DAYS_IN_MONTH
    ramp_roles = {"SMB Account Executive", "ENT Account Executive",
                  "SMB Account Manager", "ENT Account Manager"}
    ramp_pct = {  # (month1, month2)
        "SMB Account Executive": (1.0, 0.5), "ENT Account Executive": (1.0, 0.5),
        "SMB Account Manager": (1.0, 1.0), "ENT Account Manager": (1.0, 1.0)}

    def ramp_credit(row):
        if row.role not in ramp_roles or pd.isna(row.role_start):
            return 0.0, False
        tot, any_ramp = 0.0, False
        for dt in days:
            if not employed(row, dt):
                continue
            n = (dt - row.role_start).days + 1
            if 1 <= n <= 30:
                pct = ramp_pct[row.role][0]
            elif 31 <= n <= 60:
                pct = ramp_pct[row.role][1]
            else:
                continue
            any_ramp = True
            tot += pct * row.quota / DAYS_IN_MONTH
        return tot, any_ramp

    rc = p.apply(ramp_credit, axis=1, result_type="expand")
    p["ramp_credit"], p["ramping"] = rc[0], rc[1]
    by_crm = p.set_index("crm_owner_id")

    # ---------- accounts ----------
    o = o.copy()
    won = o.opportunity_stage.eq("Closed Won") if stage_rule == "stage" else o.closed_won_date.notna()
    cw = o[won]
    if go_live == "earliest_cw":
        gl = cw.groupby("crm_account_id").effective_start_date.min()
    else:  # naive: latest effective start on any Closed Won
        gl = cw.groupby("crm_account_id").effective_start_date.max()
    first_opp = o.sort_values("effective_start_date").groupby("crm_account_id").first()
    acct = first_opp[["closed_by_owner_id", "crm_owner_id", "account_type", "segment",
                      "charge_rate_per_shift", "pay_rate_per_shift", "net_terms"]].copy()
    acct["go_live"] = gl
    acct = acct.reset_index().merge(m.drop(columns="src_row"), on="crm_account_id")
    acct["gtm"] = acct.account_type.eq("GTM") & ~acct.crm_account_id.isin(gtm_exclude)
    acct["nr_per_shift"] = acct.charge_rate_per_shift - acct.pay_rate_per_shift
    acct["terms_mult"] = acct.net_terms.map(TERMS_MULT)

    s = a.merge(acct, on="app_account_id")
    s["F"] = s.account_first_worked_shift_date
    s["nr"] = s.nr_per_shift.where(s.worked_shift, 0.0)
    s["gsv"] = s.charge_rate_per_shift
    s["wdate"] = s.shift_worked_date
    s["pdate"] = s.shift_posted_date

    def window(dt, F):
        n = (dt - F).dt.days + 1
        return pd.cut(n, [-10**6, 30, 60, 90, 10**6], labels=["W1", "W2", "W3", "after"])

    # NR windows by worked date (non-GTM); GSV windows by posted date (GTM, upper bound only)
    s["win_nr"] = window(s.wdate, s.F)
    s["win_gsv"] = window(s.pdate, s.F)

    # ---------- AE new-logo credit ----------
    rows = []
    for acc_id, g in s.groupby("crm_account_id"):
        A = g.iloc[0]
        if pd.isna(A.closed_by_owner_id):
            closer = None
        else:
            closer = A.closed_by_owner_id
        F = A.F
        if A.gtm:
            val, win, dt = g.gsv, g.win_gsv, g.pdate
        else:
            val, win, dt = g.nr.where(g.worked_shift, 0.0), g.win_nr, g.wdate
        W = {w: val[win == w].sum() for w in ["W1", "W2", "W3"]}
        mult = 1.0 if A.gtm else A.terms_mult
        # F1-30D: June-recognised dollars in W1
        j = (win == "W1") & dt.between(MONTH_START, MONTH_END)
        w1_june = val[j].sum()
        if A.gtm and catch_up_gtm:
            # A9: W1 GSV posted before June but never paid (window unknown until F) -> June, if F in June
            if MONTH_START <= F <= MONTH_END:
                w1_june += val[(win == "W1") & (dt < MONTH_START)].sum()
        for comp_name, amt, rec_date, half in [
            ("F1-30D", w1_june, None, False),
            ("F31-60D", max(0.0, W["W2"] - W["W1"]), F + pd.Timedelta(days=59), True),
            ("F61-90D", max(0.0, W["W3"] - max(W["W1"], W["W2"])), F + pd.Timedelta(days=89), True)]:
            if comp_name != "F1-30D" and not (MONTH_START <= rec_date <= MONTH_END):
                continue
            if amt == 0:
                continue
            rows.append(dict(crm_account_id=acc_id, closer=closer, gtm=A.gtm, component=comp_name,
                             raw=amt, mult=mult, half=half, rec_date=rec_date,
                             june_mask_dates=None))
    ae = pd.DataFrame(rows)

    # Employment test on recognition date: F1-30D per shift date -> need shift-level split
    def f130_split(acc_id, closer_row):
        g = s[s.crm_account_id == acc_id]
        A = g.iloc[0]
        if A.gtm:
            val, win, dt = g.gsv, g.win_gsv, g.pdate
        else:
            val, win, dt = g.nr, g.win_nr, g.wdate
        j = (win == "W1") & dt.between(MONTH_START, MONTH_END)
        if A.gtm and catch_up_gtm and MONTH_START <= A.F <= MONTH_END:
            pre = (win == "W1") & (dt < MONTH_START)
            # catch-up recognised on the first worked shift date
            dts = pd.concat([dt[j], pd.Series(A.F, index=dt[pre].index)])
            vals = pd.concat([val[j], val[pre]])
        else:
            dts, vals = dt[j], val[j]
        if closer_row is None:
            return 0.0, vals.sum()
        ok = dts.apply(lambda x: employed(closer_row, x))
        return vals[ok].sum(), vals[~ok].sum()

    out = []
    for _, row in ae.iterrows():
        cr_row = by_crm.loc[row.closer] if row.closer in by_crm.index else None
        factor = row.mult * (HALF if row.half else 1.0)
        if row.gtm:
            role_rate = cr_row.rate if cr_row is not None else 0.16
            factor *= GTM_RATE / role_rate
        if row.component == "F1-30D":
            cred, uncred = f130_split(row.crm_account_id, cr_row)
        else:
            if cr_row is not None and employed(cr_row, row.rec_date):
                cred, uncred = row.raw, 0.0
            else:
                cred, uncred = 0.0, row.raw
        out.append({**row, "credited": cred * factor, "uncredited": uncred * factor})
    ae = pd.DataFrame(out)

    # ---------- AM existing credit ----------
    existing = acct[acct.go_live < MONTH_START].crm_account_id
    sj = s[s.crm_account_id.isin(existing) & s.wdate.between(MONTH_START, MONTH_END)]
    am_nr = sj.groupby("crm_owner_id").nr.sum()
    bb = b[(b.ticket_status == "approved")].merge(acct[["billing_account_id", "crm_account_id", "crm_owner_id", "go_live"]],
                                                   on="billing_account_id")
    bj = bb[(bb.billing_month == "2026-06") & (bb.go_live < MONTH_START)]
    am_cr = bj.groupby("crm_owner_id").credit_amount.sum()
    # May, for SPIFF: book held in June that went live before 1 May
    book_may = acct[acct.go_live < pd.Timestamp("2026-05-01")]
    def ecnr(month_start, month_end, bm, accts):
        x = s[s.crm_account_id.isin(accts.crm_account_id) & s.wdate.between(month_start, month_end)].groupby("crm_owner_id").nr.sum()
        c = bb[(bb.billing_month == bm) & bb.crm_account_id.isin(accts.crm_account_id)].groupby("crm_owner_id").credit_amount.sum()
        return x.sub(c, fill_value=0)
    spiff_june = ecnr(MONTH_START, MONTH_END, "2026-06", book_may)
    spiff_may = ecnr(pd.Timestamp("2026-05-01"), pd.Timestamp("2026-05-31"), "2026-05", book_may)

    # ---------- per-rep ----------
    acc_t = acc.copy()
    res = p[["name", "hr_id", "crm_owner_id", "role", "Country Code", "Manager - ID", "quota",
             "quota_prorated", "rate", "ramp_credit", "ramping", "days_employed", "emp_end"]].copy()
    ae_by = ae.groupby("closer")[["credited", "uncredited"]].sum()
    res["new_produced"] = res.crm_owner_id.map(ae_by.credited).fillna(0.0)
    res["new_uncredited"] = res.crm_owner_id.map(ae_by.uncredited).fillna(0.0)
    res["exist_nr"] = res.crm_owner_id.map(am_nr).fillna(0.0)
    res["exist_credits"] = res.crm_owner_id.map(am_cr).fillna(0.0)
    res["exist_produced"] = res.exist_nr - res.exist_credits
    is_ae = res.role.str.contains("Account Executive")
    is_am = res.role.isin(["SMB Account Manager", "ENT Account Manager"])
    res["produced"] = 0.0
    res.loc[is_ae, "produced"] = res.new_produced
    res.loc[is_am, "produced"] = res.exist_produced
    res["credited"] = res.produced
    ramp_on = res.ramping & (is_ae | is_am)
    res.loc[ramp_on, "credited"] = res.loc[ramp_on, ["produced", "ramp_credit"]].max(axis=1)

    bands = {"SMB Account Executive": (30000, 40000, *accel_smb),
             "ENT Account Executive": (60000, 80000, 1.2, 1.3)}

    def commission(row):
        c = row.credited
        if row.role in bands and not row.ramping:
            t1, t2, m1, m2 = bands[row.role]
            std = min(c, t1); b1 = max(0, min(c, t2) - t1); b2 = max(0, c - t2)
            return row.rate * (std + m1 * b1 + m2 * b2)
        if c < 0 and negative_am == "floor0":
            return 0.0
        return row.rate * c
    res["commission_rev"] = res.apply(commission, axis=1)

    spiff_amt = {"SMB Account Manager": 500.0, "ENT Account Manager": 1500.0}
    res["may_book"] = res.crm_owner_id.map(spiff_may).fillna(0.0)
    res["june_book"] = res.crm_owner_id.map(spiff_june).fillna(0.0)
    res["spiff"] = 0.0
    ok = is_am & ~res.ramping & (res.may_book > 0) & (res.june_book >= res.may_book)
    res.loc[ok, "spiff"] = res.loc[ok, "role"].map(spiff_amt)

    # ---------- managers ----------
    hr2name = dict(zip(res.hr_id, res.name))
    parent = dict(zip(res.hr_id, res["Manager - ID"]))
    def ancestors(hid):
        out, cur = [], parent.get(hid)
        while cur in hr2name:
            out.append(cur); cur = parent.get(cur)
        return out
    carry_new = {hid: 0.0 for hid in res.hr_id}
    carry_ex = {hid: 0.0 for hid in res.hr_id}
    for _, row in res.iterrows():
        nv = (row.credited + row.new_uncredited) if row.role.endswith("Account Executive") else 0.0
        ev = row.credited if row.role in ("SMB Account Manager", "ENT Account Manager") else 0.0
        for anc in ancestors(row.hr_id):
            carry_new[anc] += nv; carry_ex[anc] += ev
    res["carry_new"] = res.hr_id.map(carry_new); res["carry_ex"] = res.hr_id.map(carry_ex)
    mgr = ~(is_ae | is_am)
    new_mgr = res.role.str.contains("AE Manager") | res.role.eq("Director of Sales")
    ex_mgr = res.role.str.contains("AM Manager") | res.role.isin(["Director of SMB AM", "Director of ENT AM"])
    res.loc[new_mgr, "credited"] = res.loc[new_mgr, "carry_new"]
    res.loc[ex_mgr, "credited"] = res.loc[ex_mgr, "carry_ex"]
    res.loc[new_mgr | ex_mgr, "commission_rev"] = res.loc[new_mgr | ex_mgr].apply(lambda x: x.rate * x.credited, axis=1)
    head = res.role.eq("Head of Sales & AM")
    rate2 = float(cr.loc["Head of Sales & AM", "Standard Payout 2"])
    res.loc[head, "commission_rev"] = res.loc[head].apply(lambda x: x.rate * x.carry_new + rate2 * x.carry_ex, axis=1)
    res.loc[head, "credited"] = res.loc[head, "carry_new"]

    res["total_usd"] = (res.commission_rev + res.spiff).round(2)
    fxr = fx.set_index("Country Code")["1 USD ="]
    res["fx"] = res["Country Code"].map(fxr)
    res["total_local"] = (res.total_usd * res.fx).round(2)
    res["attainment"] = res.credited / res.quota_prorated.where(res.quota_prorated > 0)
    return dict(people=res, ae=ae, acct=acct, shifts=s)


if __name__ == "__main__":
    pd.set_option("display.width", 250); pd.set_option("display.max_rows", 200)
    out = build()
    p = out["people"]
    print(p[["name", "role", "produced", "ramp_credit", "credited", "quota_prorated", "attainment",
             "commission_rev", "spiff", "total_usd", "total_local"]].round(2).to_string())
    print("TOTAL USD", round(p.total_usd.sum(), 2))
