"""
Honeywell Binstock Program Scorecard — Dashboard Generator (Tempe Site)
-----------------------------------------------------------------------
Run:    python generate_dashboard_tempe.py
Output: index.html  (saved to the same folder as the xlsx)

Requirements:
    pip install pandas openpyxl
"""

import pandas as pd
import os
from datetime import datetime

# ── CONFIG ───────────────────────────────────────────────────────────────────
XLSX_PATH  = r"C:\Users\zn424f\OneDrive - The Boeing Company\Working KPIs\Bin Stratifications\Honeywell Tempe\Honeywell Tempe_042026_binstrat.xlsx"
SHEET_NAME = "Bin Map Rpt_Tempe"
OUTPUT_FILE = "index.html"
# ─────────────────────────────────────────────────────────────────────────────


def load_and_calculate(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet)

    total    = len(df)
    active   = len(df[df["Bin Activity Status"] == "Active"])
    inactive = len(df[df["Bin Activity Status"] == "Inactive"])

    stockout_total  = len(df[df["Stockout Status"] == "STOCKOUT"])
    stockout_active = len(df[(df["Stockout Status"] == "STOCKOUT") & (df["Bin Activity Status"] == "Active")])

    fill_total  = round((total  - stockout_total)  / total  * 100, 2) if total  else 0
    fill_active = round((active - stockout_active) / active * 100, 2) if active else 0

    past_due_total  = len(df[df["Past Due?"] == "Yes"])
    past_due_active = len(df[(df["Past Due?"] == "Yes") & (df["Bin Activity Status"] == "Active")])

    pd_pct_total  = round(past_due_total  / total  * 100, 2) if total  else 0
    pd_pct_active = round(past_due_active / active * 100, 2) if active else 0

    on_contract  = len(df[df["Contract Status"] == "On-Contract Priced"])
    off_contract = len(df[df["Contract Status"] == "Off-Contract"])
    unpriced     = len(df[df["Contract Status"] == "Unpriced"])

    on_contract_pct  = round(on_contract  / total * 100, 2) if total else 0
    off_contract_pct = round(off_contract / total * 100, 2) if total else 0
    unpriced_pct     = round(unpriced     / total * 100, 2) if total else 0

    active_pct   = round(active   / total * 100, 2) if total else 0
    inactive_pct = round(inactive / total * 100, 2) if total else 0

    stockout_pct_active = round(stockout_active / active * 100, 2) if active else 0

    return {
        "report_date"        : datetime.now().strftime("%B %d, %Y"),
        "file_name"          : os.path.basename(path),
        "total"              : f"{total:,}",
        "active"             : f"{active:,}",
        "inactive"           : f"{inactive:,}",
        "active_pct"         : active_pct,
        "inactive_pct"       : inactive_pct,
        "fill_total"         : fill_total,
        "fill_active"        : fill_active,
        "stockout_total"     : stockout_total,
        "stockout_active"    : stockout_active,
        "stockout_pct_active": stockout_pct_active,
        "past_due_total"     : past_due_total,
        "past_due_active"    : past_due_active,
        "pd_pct_total"       : pd_pct_total,
        "pd_pct_active"      : pd_pct_active,
        "on_contract"        : on_contract,
        "off_contract"       : off_contract,
        "unpriced"           : unpriced,
        "on_contract_pct"    : on_contract_pct,
        "off_contract_pct"   : off_contract_pct,
        "unpriced_pct"       : unpriced_pct,
    }


def donut_path(pct, r=52, cx=64, cy=64):
    """Return SVG arc path and dasharray for a donut segment."""
    circ = 2 * 3.14159 * r
    filled = circ * (pct / 100)
    return f"{filled:.2f} {circ - filled:.2f}"


def build_html(d):
    fill_total_arc  = donut_path(d["fill_total"])
    fill_active_arc = donut_path(d["fill_active"])

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Honeywell Binstock Program Scorecard — Tempe</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&family=Syne:wght@700;800&display=swap" rel="stylesheet"/>
<style>
  /* ── TOKENS ─────────────────────────────── */
  :root {{
    --bg         : #f5f6f8;
    --surface    : #ffffff;
    --border     : #d4d8de;
    --text       : #1a1d23;
    --subtext    : #4a5060;
    --muted      : #6b7280;
    --accent     : #1d4ed8;
    --accent-lt  : #dbeafe;
    --green      : #15803d;
    --green-lt   : #dcfce7;
    --red        : #b91c1c;
    --red-lt     : #fee2e2;
    --yellow     : #92400e;
    --yellow-lt  : #fef3c7;
    --radius     : 12px;
    --shadow     : 0 1px 4px rgba(0,0,0,.08);
  }}

  /* ── RESET ───────────────────────────────── */
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Inter', system-ui, sans-serif;
    font-size: 14px;
    font-weight: 600;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    padding: 24px 20px 40px;
  }}

  /* ── LAYOUT ──────────────────────────────── */
  .page       {{ max-width: 1100px; margin: 0 auto; }}
  .grid-2     {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
  .grid-3     {{ display: grid; grid-template-columns: repeat(3,1fr); gap: 16px; }}
  .grid-4     {{ display: grid; grid-template-columns: repeat(4,1fr); gap: 16px; }}
  .col-span-2 {{ grid-column: span 2; }}
  .col-span-3 {{ grid-column: span 3; }}
  .col-span-4 {{ grid-column: span 4; }}
  .mt16       {{ margin-top: 16px; }}

  /* ── CARD ────────────────────────────────── */
  .card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    box-shadow: var(--shadow);
    padding: 20px;
  }}
  .card-title {{
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .06em;
    text-transform: uppercase;
    color: var(--muted);
    margin-bottom: 14px;
  }}

  /* ── HEADER ──────────────────────────────── */
  .header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 24px;
    padding: 18px 24px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    box-shadow: var(--shadow);
  }}
  .header-left h1 {{
    font-family: 'Syne', sans-serif;
    font-weight: 800;
    font-size: 22px;
    color: var(--text);
    letter-spacing: -.02em;
  }}
  .header-left .sub {{
    font-size: 12px;
    font-weight: 600;
    color: var(--subtext);
    margin-top: 3px;
  }}
  .header-right {{
    text-align: right;
    font-size: 12px;
    font-weight: 600;
    color: var(--subtext);
    line-height: 1.6;
  }}
  .site-badge {{
    display: inline-block;
    background: var(--accent-lt);
    color: var(--accent);
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .05em;
    text-transform: uppercase;
    padding: 2px 10px;
    border-radius: 20px;
    margin-top: 6px;
  }}

  /* ── KPI STAT ────────────────────────────── */
  .stat-val {{
    font-family: 'Syne', sans-serif;
    font-size: 32px;
    font-weight: 800;
    color: var(--text);
    line-height: 1;
  }}
  .stat-label {{
    font-size: 12px;
    font-weight: 700;
    color: var(--subtext);
    margin-top: 4px;
  }}
  .stat-sub {{
    font-size: 12px;
    font-weight: 600;
    color: var(--muted);
    margin-top: 6px;
  }}

  /* ── DONUT ───────────────────────────────── */
  .donut-wrap {{
    display: flex;
    align-items: center;
    gap: 20px;
  }}
  .donut-svg {{ flex-shrink: 0; }}
  .donut-track  {{ fill: none; stroke: #e5e7eb; stroke-width: 10; }}
  .donut-fill   {{ fill: none; stroke-width: 10; stroke-linecap: round;
                   stroke-dashoffset: 0; transform-origin: center;
                   transform: rotate(-90deg); transition: stroke-dasharray .8s ease; }}
  .donut-green  {{ stroke: #16a34a; }}
  .donut-blue   {{ stroke: #2563eb; }}
  .donut-center {{ font-family: 'Syne', sans-serif; font-size: 17px; font-weight: 800;
                   fill: var(--accent); text-anchor: middle; dominant-baseline: middle; }}
  .donut-legend {{ flex: 1; }}
  .donut-legend .lg-row {{ display: flex; justify-content: space-between;
                            align-items: center; margin-bottom: 8px; }}
  .donut-legend .lg-label {{ font-size: 12px; font-weight: 700; color: var(--subtext); }}
  .donut-legend .lg-val   {{ font-family: 'Syne', sans-serif; font-size: 15px;
                              font-weight: 800; color: var(--text); }}

  /* ── FILL BAR ────────────────────────────── */
  .fill-bar {{ margin-top: 14px; }}
  .fill-bar-label {{
    display: flex;
    justify-content: space-between;
    font-size: 12px;
    font-weight: 700;
    color: var(--subtext);
    margin-bottom: 5px;
  }}
  .fill-bar-track {{
    background: #e5e7eb;
    border-radius: 99px;
    height: 8px;
    overflow: hidden;
  }}
  .fill-bar-fill {{
    height: 100%;
    border-radius: 99px;
    transition: width 1s ease;
  }}
  .bar-green {{ background: #16a34a; }}
  .bar-blue  {{ background: #2563eb; }}

  /* ── ACTIVITY DOT MATRIX ─────────────────── */
  #activityDots {{
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
    margin-top: 10px;
  }}
  .dot-bin {{
    width: 10px;
    height: 10px;
    border-radius: 3px;
    flex-shrink: 0;
  }}
  .dot-active   {{ background: #16a34a; }}
  .dot-inactive {{ background: #e5e7eb; }}
  .dot-legend {{
    display: flex;
    gap: 16px;
    margin-top: 10px;
    font-size: 12px;
    font-weight: 700;
    color: var(--subtext);
    align-items: center;
  }}
  .dot-legend span {{ display: flex; align-items: center; gap: 5px; }}
  .dot-legend i {{
    display: inline-block;
    width: 10px; height: 10px;
    border-radius: 3px;
  }}

  /* ── RISK PANEL ──────────────────────────── */
  .risk-row {{
    display: grid;
    grid-template-columns: 170px 1fr 52px 80px;
    align-items: center;
    gap: 10px;
    margin-bottom: 10px;
  }}
  .risk-label {{ font-size: 12px; font-weight: 700; color: var(--subtext); }}
  .risk-bar-wrap {{ display: flex; align-items: center; }}
  .risk-bar-bg {{
    flex: 1;
    background: #e5e7eb;
    border-radius: 99px;
    height: 7px;
    overflow: hidden;
  }}
  .risk-bar-inner {{ height: 100%; border-radius: 99px; }}
  .bg-red    {{ background: #dc2626; }}
  .bg-yellow {{ background: #d97706; }}
  .bg-green  {{ background: #16a34a; }}
  .risk-value {{ font-family: 'Syne', sans-serif; font-size: 14px; font-weight: 800; text-align: right; }}
  .risk-count {{ font-size: 12px; font-weight: 600; text-align: right; }}
  .c-red    {{ color: var(--red); }}
  .c-yellow {{ color: var(--yellow); }}
  .c-green  {{ color: var(--green); }}
  .c-muted  {{ color: var(--muted); }}

  /* ── BADGE PILLS ─────────────────────────── */
  .pill {{
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .03em;
  }}
  .pill-green  {{ background: var(--green-lt);  color: var(--green);  }}
  .pill-red    {{ background: var(--red-lt);    color: var(--red);    }}
  .pill-yellow {{ background: var(--yellow-lt); color: var(--yellow); }}
  .pill-blue   {{ background: var(--accent-lt); color: var(--accent); }}

  /* ── FOOTER ──────────────────────────────── */
  .footer {{
    margin-top: 28px;
    font-size: 11px;
    font-weight: 600;
    color: var(--muted);
    display: flex;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 6px;
  }}
</style>
</head>
<body>

<div class="page">

  <!-- HEADER -->
  <div class="header">
    <div class="header-left">
      <h1>Honeywell Binstock Program Scorecard</h1>
      <div class="sub">Boeing Distribution Services · Program Management</div>
      <div class="site-badge">Tempe Site</div>
    </div>
    <div class="header-right">
      <div>{d['report_date']}</div>
      <div>{d['file_name']}</div>
    </div>
  </div>

  <!-- ROW 1: KPI STATS -->
  <div class="grid-4">

    <div class="card">
      <div class="card-title">Total Bin Map</div>
      <div class="stat-val">{d['total']}</div>
      <div class="stat-label">Managed Locations</div>
      <div class="stat-sub">
        <span class="pill pill-green">{d['active']:,} Active</span>&nbsp;
        <span class="pill pill-red">{d['inactive']:,} Inactive</span>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Stockout Exposure</div>
      <div class="stat-val c-red">{d['stockout_total']}</div>
      <div class="stat-label">Empty Bins (Total Map)</div>
      <div class="stat-sub">
        <span class="pill pill-red">{d['stockout_active']} in Active Bins</span>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Past Due Risk</div>
      <div class="stat-val c-yellow">{d['past_due_total']}</div>
      <div class="stat-label">Past Due Bins (Total Map)</div>
      <div class="stat-sub">
        <span class="pill pill-yellow">{d['past_due_active']} Active · {d['pd_pct_active']}%</span>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Contract Coverage</div>
      <div class="stat-val c-green">{d['on_contract_pct']}%</div>
      <div class="stat-label">On-Contract Priced</div>
      <div class="stat-sub">
        <span class="pill pill-yellow">{d['unpriced']} Unpriced</span>&nbsp;
        <span class="pill pill-red">{d['off_contract']} Off-Contract</span>
      </div>
    </div>

  </div><!-- /grid-4 -->

  <!-- ROW 2: FILL RATE + ACTIVITY -->
  <div class="grid-2 mt16">

    <!-- FILL RATE CARD -->
    <div class="card">
      <div class="card-title">Fill Rate — Dual Lens</div>
      <div class="grid-2">

        <div>
          <div class="donut-wrap">
            <svg class="donut-svg" width="128" height="128" viewBox="0 0 128 128">
              <circle class="donut-track" cx="64" cy="64" r="52"/>
              <circle class="donut-fill donut-green" cx="64" cy="64" r="52"
                stroke-dasharray="{fill_total_arc}"/>
              <text class="donut-center" x="64" y="64">{d['fill_total']}%</text>
            </svg>
            <div class="donut-legend">
              <div class="lg-row">
                <span class="lg-label">Filled Bins</span>
                <span class="lg-val c-green">{int(d['total'].replace(',','')) - d['stockout_total']:,}</span>
              </div>
              <div class="lg-row">
                <span class="lg-label">Stockout</span>
                <span class="lg-val c-red">{d['stockout_total']}</span>
              </div>
            </div>
          </div>
          <div class="fill-bar">
            <div class="fill-bar-label"><span>vs. Total Bin Map</span><span>{d['fill_total']}%</span></div>
            <div class="fill-bar-track"><div class="fill-bar-fill bar-green" style="width:{d['fill_total']}%;"></div></div>
          </div>
        </div>

        <div>
          <div class="donut-wrap">
            <svg class="donut-svg" width="128" height="128" viewBox="0 0 128 128">
              <circle class="donut-track" cx="64" cy="64" r="52"/>
              <circle class="donut-fill donut-blue" cx="64" cy="64" r="52"
                stroke-dasharray="{fill_active_arc}"/>
              <text class="donut-center" x="64" y="64">{d['fill_active']}%</text>
            </svg>
            <div class="donut-legend">
              <div class="lg-row">
                <span class="lg-label">Filled Active</span>
                <span class="lg-val c-green">{int(d['active'].replace(',','')) - d['stockout_active']:,}</span>
              </div>
              <div class="lg-row">
                <span class="lg-label">Stockout</span>
                <span class="lg-val c-red">{d['stockout_active']}</span>
              </div>
            </div>
          </div>
          <div class="fill-bar">
            <div class="fill-bar-label"><span>vs. Active Bins Only</span><span>{d['fill_active']}%</span></div>
            <div class="fill-bar-track"><div class="fill-bar-fill bar-blue" style="width:{d['fill_active']}%;"></div></div>
          </div>
        </div>

      </div>
    </div><!-- /fill rate -->

    <!-- BIN ACTIVITY CARD -->
    <div class="card">
      <div class="card-title">Bin Activity — 3-Year Scan History</div>

      <div style="display:flex; justify-content:space-between; margin-bottom:12px;">
        <div>
          <div class="stat-val c-green">{d['active_pct']}%</div>
          <div class="stat-label">Active Bins</div>
          <div class="stat-sub">{d['active']} locations</div>
        </div>
        <div style="text-align:right;">
          <div class="stat-val" style="color:var(--muted);">{d['inactive_pct']}%</div>
          <div class="stat-label">Inactive Bins</div>
          <div class="stat-sub">{d['inactive']} locations</div>
        </div>
      </div>

      <div id="activityDots"></div>
      <div class="dot-legend">
        <span><i style="background:#16a34a;"></i> Active (scanned 2023–2026)</span>
        <span><i style="background:#e5e7eb;"></i> Inactive (0 scans)</span>
      </div>

      <div class="fill-bar" style="margin-top:14px;">
        <div class="fill-bar-label"><span>Activity Rate</span><span>{d['active_pct']}%</span></div>
        <div class="fill-bar-track"><div class="fill-bar-fill bar-green" style="width:{d['active_pct']}%;"></div></div>
      </div>
    </div><!-- /activity -->

  </div><!-- /grid-2 -->

  <!-- ROW 3: CONTRACT RISK PANEL -->
  <div class="card mt16">
    <div class="card-title">Contract &amp; Risk Summary</div>
    <div class="grid-2">

      <div>
        <div class="risk-row">
          <div class="risk-label">On-Contract Priced</div>
          <div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-green" style="width:{d['on_contract_pct']}%;"></div></div></div>
          <div class="risk-value c-green">{d['on_contract_pct']}%</div>
          <div class="risk-count c-muted">{d['on_contract']} bins</div>
        </div>
        <div class="risk-row">
          <div class="risk-label">Off-Contract</div>
          <div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-red" style="width:{d['off_contract_pct']}%;"></div></div></div>
          <div class="risk-value c-red">{d['off_contract_pct']}%</div>
          <div class="risk-count c-muted">{d['off_contract']} bins</div>
        </div>
      </div>

      <div>
        <div class="risk-row">
          <div class="risk-label">Unpriced (Total Map)</div>
          <div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-yellow" style="width:{d['unpriced_pct']}%;"></div></div></div>
          <div class="risk-value c-yellow">{d['unpriced_pct']}%</div>
          <div class="risk-count c-muted">{d['unpriced']} bins</div>
        </div>
        <div class="risk-row">
          <div class="risk-label">Past Due (Active Lens)</div>
          <div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-red" style="width:{d['pd_pct_active']}%;"></div></div></div>
          <div class="risk-value c-red">{d['pd_pct_active']}%</div>
          <div class="risk-count c-muted">{d['past_due_active']} bins</div>
        </div>
        <div class="risk-row">
          <div class="risk-label">Stockout (Active Lens)</div>
          <div class="risk-bar-wrap"><div class="risk-bar-bg"><div class="risk-bar-inner bg-red" style="width:{d['stockout_pct_active']}%;"></div></div></div>
          <div class="risk-value c-red">{d['stockout_pct_active']}%</div>
          <div class="risk-count c-muted">{d['stockout_active']} bins</div>
        </div>
      </div>

    </div>
  </div><!-- /risk panel -->

</div><!-- /page -->

<div class="footer">
  <span>Boeing Distribution Services · Program Management · Honeywell Aerospace Account</span>
  <span>Data: {d['file_name']} · {d['report_date']} · Active = any scan 2023–2026 · Inactive = 0 scans same period</span>
</div>

<script>
  const container = document.getElementById('activityDots');
  const activeDots = Math.round(({d['active_pct']} / 100) * 150);
  for (let i = 0; i < 150; i++) {{
    const dot = document.createElement('div');
    dot.className = 'dot-bin ' + (i < activeDots ? 'dot-active' : 'dot-inactive');
    container.appendChild(dot);
  }}
  window.addEventListener('load', () => {{
    document.querySelectorAll('.fill-bar-fill').forEach(el => {{
      const w = el.style.width;
      el.style.width = '0%';
      requestAnimationFrame(() => {{ setTimeout(() => {{ el.style.width = w; }}, 100); }});
    }});
  }});
</script>

</body>
</html>"""


if __name__ == "__main__":
    print(f"Reading: {XLSX_PATH}")
    data = load_and_calculate(XLSX_PATH, SHEET_NAME)
    html  = build_html(data)
    output_path = os.path.join(os.path.dirname(XLSX_PATH), OUTPUT_FILE)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✓ Dashboard generated: {output_path}")
