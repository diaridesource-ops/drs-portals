#!/usr/bin/env python3
"""
DRS Management Statement — manager-facing fleet view (Lindsey only).

Renders one HTML page from turo.xlsx with, for every vehicle in the fleet:
current statement month payout, revenue, DRS management fee, expenses,
upcoming booked revenue — plus fleet totals and prior-month recaps
(6/2026 forward, same data window as the owner portals).

Used by build_all.py (GitHub Actions) to publish to the _mgmt slug.
Can also be run directly:  python3 mgmt_statement.py [--out FILE]
"""
import sys, datetime, html
import portal_generator as pg


def month_list(today):
    cur = (today.year, today.month)
    months, y, m = [], *pg.DATA_START
    while (y, m) <= cur:
        months.append((y, m))
        m += 1
        if m == 13:
            m, y = 1, y + 1
    return months  # oldest -> current


def render_mgmt(today):
    roster, stmts, exps, trips = pg.load(today)
    cur = (today.year, today.month)
    months = month_list(today)

    # Build one record per VIN
    vins = sorted(roster.keys(), key=lambda v: (roster[v]['owner'].lower(), v))

    def stmt_for(vin, month):
        return next((s for s in stmts.get(vin, []) if s['month'] == month), None)

    def exp_total(vin, month):
        return sum(e['amount'] for e in exps.get(vin, []) if e['month'] == month)

    # ---- fleet KPIs (current month) ----
    tot = {'revenue': 0.0, 'fee': 0.0, 'payout': 0.0, 'days': 0.0, 'upcoming': 0.0, 'trips': 0}
    for vin in vins:
        s = stmt_for(vin, cur)
        if s:
            tot['revenue'] += s['revenue']; tot['fee'] += s['mgmt_fee']
            tot['payout'] += s['payout']; tot['days'] += s['days']
        tot['upcoming'] += sum(t['rss'] for t in trips.get(vin, []))
        tot['trips'] += len(trips.get(vin, []))

    # ---- data alerts: activity that doesn't match any roster vehicle ----
    orphans = sorted((set(stmts) | set(exps) | set(trips)) - set(roster))
    alerts = []
    for vin in orphans:
        srows = stmts.get(vin, []); erows = exps.get(vin, []); trows = trips.get(vin, [])
        who = srows[0]['owner'] if srows else (erows[0]['vendor'] if erows else '?')
        amt = sum(s['payout'] for s in srows) + sum(-e['amount'] for e in erows if e['credit'])
        label = f'VIN "{vin}"' if vin else 'BLANK VIN'
        alerts.append(
            f"<li><b>{label}</b> ({html.escape(str(who))}): {len(srows)} statement rows, "
            f"{len(erows)} expenses, {len(trows)} trips are not linked to any vehicle in the "
            f"Owners tab — {pg.money(amt)} affected. Fix the VIN in the sheet so this data "
            f"reaches the owner's portal and these fleet totals.</li>")
    alert_html = (
        '<div class="note amber"><b>&#9888; Data alerts — fix in Turo Master:</b><ul>'
        + ''.join(alerts) + '</ul><span class="fine">Everything below covers roster '
        'vehicles only; the amounts above are excluded until the VINs match.</span></div>'
    ) if alerts else ''

    kpis = f"""
  <div class="kpis">
    <div class="kpi"><div class="klabel">Fleet revenue ({pg.mlabel(cur)})</div>
      <div class="kval">{pg.money(tot['revenue'])}</div></div>
    <div class="kpi"><div class="klabel">DRS management fees ({pg.mlabel(cur)})</div>
      <div class="kval pos">{pg.money(tot['fee'])}</div></div>
    <div class="kpi"><div class="klabel">Owner payouts due ({pg.mlabel(cur)})</div>
      <div class="kval {'neg' if tot['payout'] < 0 else ''}">{pg.money(tot['payout'])}</div></div>
    <div class="kpi"><div class="klabel">Upcoming booked revenue</div>
      <div class="kval">{pg.money(tot['upcoming'])}</div>
      <div class="ksub">{tot['trips']} booked trips (subject to split)</div></div>
  </div>"""

    # ---- per-vehicle table for one month ----
    def month_table(month, current=False):
        rows, t = [], {'revenue': 0.0, 'fee': 0.0, 'adj': 0.0, 'opex': 0.0, 'payout': 0.0, 'exp': 0.0}
        for vin in vins:
            info = roster[vin]
            s = stmt_for(vin, month)
            ex = exp_total(vin, month)
            veh = pg.fix_vehicle(s['vehicle'] if s else info['desc'], vin)
            if s:
                t['revenue'] += s['revenue']; t['fee'] += s['mgmt_fee']; t['adj'] += s['adjustments']
                t['opex'] += s['op_expenses']; t['payout'] += s['payout']
            t['exp'] += ex
            up = sum(tr['rss'] for tr in trips.get(vin, [])) if current else None
            pay = s['payout'] if s else None
            pc = 'r neg' if (pay is not None and pay < 0) else 'r pos'
            rows.append(
                f"<tr><td>{html.escape(info['owner'])}</td>"
                f"<td>{html.escape(veh)}<div class='fine'>{vin} · {int(info['split']*100)}/{int((1-info['split'])*100)} split</div></td>"
                f"<td class='r'>{int(s['days']) if s else '—'}</td>"
                f"<td class='r'>{pg.money(s['revenue']) if s else '—'}</td>"
                f"<td class='r'>{pg.money(s['mgmt_fee']) if s else '—'}</td>"
                f"<td class='r'>{pg.money(s['adjustments']) if s else '—'}</td>"
                f"<td class='r'>{pg.money(s['op_expenses']) if s else '—'}</td>"
                f"<td class='r'>{pg.money(ex)}</td>"
                f"<td class='{pc}'><b>{pg.money(pay) if pay is not None else '—'}</b></td>"
                + (f"<td class='r'>{pg.money(up)}<div class='fine'>{len(trips.get(vin, []))} trips</div></td>" if current else '')
                + "</tr>")
        cols = 10 if current else 9
        head = ('<tr><th>Owner</th><th>Vehicle</th><th class="r">Days</th><th class="r">Revenue</th>'
                '<th class="r">DRS Fee</th><th class="r">Adjustments</th><th class="r">Op Expenses</th>'
                '<th class="r">Expenses Charged</th><th class="r">Owner Payout</th>'
                + ('<th class="r">Upcoming Booked</th>' if current else '') + '</tr>')
        foot = (f"<tr class='total'><td colspan='3'>Fleet total — {pg.mlabel(month)}</td>"
                f"<td class='r'>{pg.money(t['revenue'])}</td><td class='r'>{pg.money(t['fee'])}</td>"
                f"<td class='r'>{pg.money(t['adj'])}</td><td class='r'>{pg.money(t['opex'])}</td>"
                f"<td class='r'>{pg.money(t['exp'])}</td>"
                f"<td class='r {'neg' if t['payout'] < 0 else ''}'><b>{pg.money(t['payout'])}</b></td>"
                + ("<td></td>" if current else '') + "</tr>")
        return f'<div class="scroll"><table>{head}<tbody>{"".join(rows)}{foot}</tbody></table></div>'

    prior = ''.join(
        f'<h2>{pg.mlabel(m)} recap</h2>{month_table(m)}'
        for m in reversed(months[:-1]))

    css_extra = """
.scroll { overflow-x:auto; }
table { min-width:900px; }
h1 .tag { font-size:12px; background:#1d4ed8; color:#fff; border-radius:6px;
  padding:2px 8px; vertical-align:middle; margin-left:8px; }
"""

    # Reuse the portal look by borrowing its CSS from a rendered page
    sample_owner = roster[vins[0]]['owner']
    portal_html = pg.render(sample_owner, roster, stmts, exps, trips, today)
    css = portal_html.split('<style>')[1].split('</style>')[0]

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>DRS Management Statement</title>
<style>{css}{css_extra}</style></head><body><div class="wrap">
<header><h1>Management Statement <span class="tag">INTERNAL</span></h1>
<div class="sub">DIA RideSource · full fleet · generated {today:%-m/%-d/%Y} · data since {pg.mlabel(pg.DATA_START)}</div></header>
{alert_html}
{kpis}
<h2>Current payouts by vehicle — {pg.mlabel(cur)}</h2>
{month_table(cur, current=True)}
{prior}
<footer>DIA RideSource, LLC · internal management statement — do not share this link with owners</footer>
</div></body></html>"""


if __name__ == '__main__':
    args = sys.argv[1:]
    out = 'mgmt.html'
    if '--out' in args:
        i = args.index('--out'); out = args[i + 1]
    today = datetime.datetime.now()
    open(out, 'w').write(render_mgmt(today))
    print('wrote', out)
