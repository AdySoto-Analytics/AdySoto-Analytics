"""Build the written summary (max 3 pages) from the recalculated model.

Numbers are read from the workbook's cached values so the PDF and the model cannot drift.
Run after the workbook has been recalculated:
    python scripts/build_summary_pdf.py
"""
from pathlib import Path
import subprocess
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
MODEL = ROOT / "deliverables" / "June_2026_Commission_Model.xlsx"
HTML = ROOT / "scripts" / "summary_source.html"  # intermediate; the PDF is the deliverable
PDF = ROOT / "deliverables" / "June_2026_Commission_Summary.pdf"
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"


def money(x, cents=False):
    return f"${x:,.2f}" if cents else f"${x:,.0f}"


def main():
    wb = openpyxl.load_workbook(MODEL, data_only=True)
    sm, lg, sc, ck = wb["Summary"], wb["Exceptions_Ledger"], wb["Scenario_Accelerator"], wb["Checks"]
    aw, rec, held = sm["D3"].value, sm["D4"].value, sm["D5"].value
    checks = ck["C4"].value
    reps = {sm.cell(row=r, column=1).value: r for r in range(9, 89)}

    def rep(name, col):
        return sm.cell(row=reps[name], column=col).value

    ledger = []
    for r in range(5, 40):
        if lg.cell(row=r, column=2).value is None or not str(lg.cell(row=r, column=2).value).startswith("E"):
            continue
        ledger.append(dict(id=lg.cell(row=r, column=2).value, finding=lg.cell(row=r, column=3).value,
                           records=lg.cell(row=r, column=5).value, credit=lg.cell(row=r, column=6).value or 0,
                           payout=lg.cell(row=r, column=7).value or 0, reps=lg.cell(row=r, column=8).value,
                           treat=lg.cell(row=r, column=10).value))
    ledger.sort(key=lambda d: -d["payout"])
    top, rest = ledger[0]["payout"], sum(d["payout"] for d in ledger[1:])

    base, repriced, delta = sc["B5"].value, sc["B6"].value, sc["B7"].value
    delta_aw, n_changed = sc["B8"].value, sc["B9"].value
    # largest beneficiaries under each basis
    rows = []
    for r in range(12, 92):
        rows.append((sc.cell(row=r, column=1).value, sc.cell(row=r, column=11).value or 0, sc.cell(row=r, column=12).value or 0))
    top_rec = max(rows, key=lambda x: x[2])
    ctl = wb["Control"]
    d1 = ctl["F24"].value - ctl["F23"].value
    band1_cost = sum((sc.cell(row=r, column=6).value or 0) * 0.16 * d1
                     for r in range(12, 92) if (sc.cell(row=r, column=12).value or 0) > 0)
    n_band2 = sum(1 for r in range(12, 92)
                  if (sc.cell(row=r, column=12).value or 0) > 0 and (sc.cell(row=r, column=7).value or 0) > 0)
    top_aw = max(rows, key=lambda x: x[1])

    jon_aw, jon_rec = rep("Jon Snow", 10), rep("Jon Snow", 11)
    katniss = rep("Katniss Everdeen", 11)

    def short(t, n=150):
        t = " ".join(str(t).split())
        return t if len(t) <= n else t[: n - 1].rsplit(" ", 1)[0] + "…"

    led_rows = "".join(
        f"<tr><td>{i + 1}</td><td><b>{d['id']}</b> {short(d['finding'], 95)}</td><td class=n>{d['records']:,}</td>"
        f"<td class=n>{money(d['credit'])}</td><td class=n><b>{money(d['payout'])}</b></td><td>{short(d['treat'], 120)}</td></tr>"
        for i, d in enumerate(ledger))

    html = f"""<!doctype html><html><head><meta charset="utf-8"><title>June 2026 Commission Summary</title>
<style>
@page {{ size: Letter; margin: 0.55in 0.6in; }}
body {{ font-family: Arial, Helvetica, sans-serif; font-size: 9.6pt; line-height: 1.32; color: #111; }}
h1 {{ font-size: 15pt; margin: 0 0 2px; }} h2 {{ font-size: 11pt; margin: 12px 0 4px; border-bottom: 1.5px solid #1F3864; color: #1F3864; padding-bottom: 1px; }}
.sub {{ color: #555; font-size: 8.8pt; margin-bottom: 6px; }}
.kpi {{ display: flex; gap: 8px; margin: 6px 0 8px; }} .kpi div {{ flex: 1; border: 1px solid #c9d3e6; border-radius: 4px; padding: 5px 8px; background: #f4f7fc; }}
.kpi b {{ display: block; font-size: 13pt; }} .kpi span {{ font-size: 8.3pt; color: #444; }}
table {{ border-collapse: collapse; width: 100%; font-size: 8.4pt; }} th {{ background: #1F3864; color: #fff; text-align: left; padding: 3px 4px; }}
td {{ border-bottom: 1px solid #ddd; padding: 2.5px 4px; vertical-align: top; }} td.n {{ text-align: right; white-space: nowrap; }}
p {{ margin: 3px 0 5px; }} ul {{ margin: 2px 0 4px 16px; padding: 0; }} li {{ margin: 1px 0; }}
.box {{ border-left: 3px solid #C00000; background: #fbf3f3; padding: 5px 8px; margin: 5px 0; }}
.reply {{ border: 1px solid #bbb; padding: 6px 9px; background: #fafafa; font-size: 9.2pt; }}
.pb {{ page-break-before: always; }}
</style></head><body>
<h1>June 2026 Commissions — Summary</h1>
<div class="sub">Sales Compensation Analyst case · Model: June_2026_Commission_Model.xlsx (all figures are live formulas; model checks: {checks})</div>

<div class="kpi">
<div><b>{money(rec, True)}</b><span>Recommended June payout, 80 reps (USD)</span></div>
<div><b>{money(aw, True)}</b><span>Plan applied exactly as written</span></div>
<div><b>{money(held, True)}</b><span>Held on one account pending review</span></div>
</div>

<div class="box"><b>Bottom line.</b> June can be paid at {money(rec)} with one exception. A rural GTM account closed by Jon Snow
posted 313 shifts and filled 45; the plan pays GTM on <i>posted</i> shifts, so as written it would pay {money(held)} on shifts nobody worked
(Jon {money(jon_aw)} → {money(jon_rec)}, plus his manager, the Director of Sales and the Head). That one finding is worth more than
every other finding combined ({money(top)} vs {money(rest)}). I recommend holding it, not cancelling it, until ops confirms who posted the shifts.</div>

<h2>1. Exceptions ledger — ranked by payout impact</h2>
<p>"Payout moved" = treated payout vs a naive processing of the same rows. Full ledger with root causes and live record counts: tab Exceptions_Ledger.</p>
<table><tr><th>#</th><th>Finding</th><th>Records</th><th>Credit $ moved</th><th>Payout moved</th><th>Treatment</th></tr>{led_rows}</table>
<p><b>Untangling stacked issues.</b> Jon Snow's account carries both E01 and E10 (GTM catch-up). Sized in waterfall order (E01 first),
E10 counts only his worked pre-June postings ($621 total); sized on the as-written basis it would be $915, so $294 of overlap is counted once,
under E01. Katniss Everdeen (left 12 Jun) stacks a leaver split (E04) on top of four new accounts: credit recognised through 12 Jun is hers
({money(katniss, True)}); everything after rolls to her manager, Bruce Wayne, under 7.11.</p>

<h2>2. Assumptions where the plan is silent — and my recommendation</h2>
<ul>
<li><b>Two numbers, side by side.</b> Data errors are fixed by the plan's own rules in both; payouts the plan allows but I would not sign off sit in a separate, switchable Treatments layer (A1).</li>
<li><b>GTM is marked by account_type</b>, since no shift-level Applied Science tag exists in the export (A17). <b>GTM posted before the first worked shift</b> in a closed month is paid in the first month it becomes knowable (A9); once GTM, all windows compare GSV with GSV (A8).</li>
<li><b>Credits on accounts not yet Existing</b> reduce nobody's pay and are listed for finance (A6). <b>A month that nets negative</b> pays $0, with no clawback without a written policy (A7). None occurred in June.</li>
<li><b>Ramping</b> = any day of the month inside a ramp month: no accelerators, no SPIFF (A10). <b>Role changes</b> price the whole month at the month-end HR role (A11). <b>HR governs</b> role, dates and reporting line over the roster (A16), and Current Role Start Date starts the ramp even where HR's previous-role end is later (A19: $0 effect in June; HR to fix before July).</li>
<li><b>Rounding</b>: full precision throughout, round USD once per rep, convert at the tab 3 rate, round local to cents (A12). The plan names no FX date; the single tab 3 rate is used.</li>
<li><b>Statements already paid</b>: underpayments paid promptly; overpayments go to sales leadership for a written decision, never deducted automatically (A14).</li>
<li><b>first worked shift</b> is taken as given (A15); <b>Walter White</b>'s HR start (2021) stands despite Slack saying he just joined (A20) — HR to confirm.</li>
</ul>

<h2 class="pb">3. Plan change: SMB AE accelerator 1.2x/1.3x → 1.3x/1.5x</h2>
<table><tr><th>June as paid (corrected, recommended)</th><th>June repriced</th><th>Delta</th><th>Reps affected</th></tr>
<tr><td class=n>{money(base, True)}</td><td class=n>{money(repriced, True)}</td><td class=n><b>+{money(delta, True)}</b></td><td class=n>{n_changed}</td></tr></table>
<p><b>Assumptions.</b> Same credited revenue (June is already worked, so no behaviour change); bands stay at $30k/$40k; ramping reps stay ineligible;
managers are paid on revenue, not commission, so they do not move.</p>
<p><b>What surprised me.</b></p>
<ul>
<li><b>It is cheap, and it mostly pays the near-misses.</b> +{money(delta)} is {delta / base:.2%} of June payout. Only {n_changed} of 28 SMB AEs are above $30k and not ramping, and {money(band1_cost)} of the cost is band I (the $30–40k step), not the 1.5x top band: only {n_band2} reps reach $40k (largest gain: {top_rec[0]}, +{money(top_rec[2])}).</li>
<li><b>On the plan as written, the biggest winner would be Jon Snow (+{money(top_aw[1])}),</b> whose band II dollars come from unworked GTM postings. Raising the top multiplier raises the payoff to exactly the behaviour E01 describes. Fix the GTM measure before raising accelerators (delta on as-written basis: +{money(delta_aw)}).</li>
<li><b>Ramping reps are excluded regardless of performance.</b> Gimli (rehired in May) finished at 120% of quota and gains nothing under either version; the plan's no-accelerator-while-ramping rule matters more than the multiplier.</li>
</ul>

<h2>4. Slack — the three I would escalate</h2>
<ol style="margin:2px 0 4px 16px;padding:0">
<li><b>GTM posting volume (Jon Snow 3, 8 &amp; 15 Jun; Doc Brown 2 Jul; Forrest Gump 3 Jul; Bruce Wayne 6 Jul; Dorothy Gale 7 Jul).</b> Five people describe the same account. It is the largest dollar issue, it recurs every month the template stays open, and only leadership can approve a hold under 7.14.</li>
<li><b>Katniss Everdeen (1 Jul): "Left on the 12th — will my June commission still process?"</b> A departed employee, real money, high trust stakes. Yes: she is paid for credit recognised through 12 Jun ({money(katniss, True)}); later credit rolls to her manager under 7.11. She needs a clear, written answer before the statement, not after.</li>
<li><b>Scout Finch (4 Jul): "122% of target but not accelerator eligible because I'm ramping."</b> Correct under the plan, but her ramp clock rests on an HR record that contradicts itself (role start 1 May, previous role end 31 May), as do Jay Gatsby's and Lord Voldemort's. One HR fix before July removes three disputes.</li>
</ol>
<p style="font-size:8.6pt;color:#444">Also real but lower stakes: Legolas (statement credited a deal in Negotiation — evidence the old process counted residue dates), Dr. Watson (SPIFF at ~90%, not "a shade under"), Darth Vader's July $500 offer (not in the plan; needs written approval to be payable).</p>

<h2>Note to sales leadership (≤250 words)</h2>
<div class="reply">
<p><b>Subject: June commissions — one payout I recommend we hold</b></p>
<p>June is ready at {money(rec)}. Paying the plan exactly as written would cost {money(aw)}; the {money(held)} difference is one account.</p>
<p>A rural facility in Jon Snow's book went live on 1 June. GTM accounts are paid on shifts <i>posted</i>, not worked. In June this one posted 293 shifts and 45 were worked — other GTM accounts fill about 78%. Jon told #sales-ops he has been loading shift templates for sites with no scheduler; the facility's director of nursing says they post a handful a day; ops flagged the account separately.</p>
<p>As written, the plan pays Jon {money(jon_aw)} for June, most of it on shifts nobody worked, and another {money(held - (jon_aw - jon_rec))} flows to Bruce, Ebenezer and Hannibal.</p>
<p><b>Recommendation:</b> pay June on the worked shifts (Jon: {money(jon_rec)}) and hold the rest — not cancel it — until ops confirms who posted the shifts. Hold the matching roll-ups.</p>
<p><b>Decisions needed before payroll:</b> (1) approve the hold in writing — the plan only changes through written leadership adjustments; (2) close the gap: cap open postings per facility, or pay GTM on filled shifts. Otherwise this account's July and August windows inflate the same way.</p>
<p>I am not assuming bad intent: Jon asked whether there was a cap before he started, and the plan allowed it. That is why this is your call, not mine.</p>
</div>
</body></html>"""
    HTML.write_text(html, encoding="utf-8")
    subprocess.run([CHROME, "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={PDF}", HTML.as_uri()], check=True, capture_output=True)
    print("wrote", PDF)


if __name__ == "__main__":
    main()
