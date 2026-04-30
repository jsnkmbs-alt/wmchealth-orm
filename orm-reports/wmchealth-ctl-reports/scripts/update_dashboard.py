"""
update_dashboard.py
-------------------
Reads metrics JSON and generates/updates docs/index.html —
the GitHub Pages dashboard showing current month KPIs and a
download link for the PowerPoint report.

Usage:
  python scripts/update_dashboard.py --data scripts/data.json
"""

import os
import json
import argparse
import glob
from datetime import datetime


def fmt_hours(hours):
    if not hours:
        return "N/A"
    if hours < 1:
        return "<1 hr"
    elif hours < 24:
        return f"{int(round(hours))} hrs"
    else:
        return f"{hours/24:.1f} days"


def rate_color(rate):
    if rate >= 90:
        return "#2E7D32"
    elif rate >= 75:
        return "#E65100"
    else:
        return "#B71C1C"


def build_html(metrics: dict, report_filename: str) -> str:
    meta = metrics["meta"]
    net = metrics["network"]
    generated = meta.get("generated_at", "")[:10]
    month_label = meta.get("month_label", "")

    # Collect all past reports for archive
    report_files = sorted(glob.glob("docs/reports/WMCHealth_CTL_*.pptx"), reverse=True)
    archive_links = ""
    for rf in report_files:
        fname = os.path.basename(rf)
        month_str = fname.replace("WMCHealth_CTL_", "").replace(".pptx", "")
        try:
            dt = datetime.strptime(month_str, "%Y-%m")
            label = dt.strftime("%B %Y")
        except Exception:
            label = month_str
        archive_links += f'<li><a href="reports/{fname}" download>📥 {label}</a></li>\n'

    # Regional cards
    region_cards = ""
    for region in ["South", "North", "West"]:
        rm = metrics["regions"].get(region, {})
        rate = rm.get("engagement_rate", 0)
        color_map = {"South": "#2E7D32", "North": "#0D7C7C", "West": "#E65100"}
        clr = color_map.get(region, "#455A64")
        rc = rate_color(rate)
        region_cards += f"""
        <div class="region-card" style="border-top: 4px solid {clr};">
          <div class="region-name" style="color:{clr};">{region}</div>
          <div class="region-rate" style="color:{rc};">{rate}%</div>
          <div class="region-sub">{rm.get('total_tasks',0)} tasks &nbsp;|&nbsp; Avg close: {fmt_hours(rm.get('avg_time_to_close_hours',0))}</div>
          <div class="sentiment-bar">
            <div class="bar-pos" style="width:{rm.get('positive',0)/(rm.get('total_tasks',1))*100:.0f}%"></div>
            <div class="bar-neu" style="width:{rm.get('neutral',0)/(rm.get('total_tasks',1))*100:.0f}%"></div>
            <div class="bar-neg" style="width:{rm.get('negative',0)/(rm.get('total_tasks',1))*100:.0f}%"></div>
          </div>
          <div class="sentiment-legend">
            <span style="color:#2E7D32;">▲ {rm.get('positive',0)} positive</span>
            <span style="color:#E65100;">◆ {rm.get('neutral',0)} neutral</span>
            <span style="color:#B71C1C;">▼ {rm.get('negative',0)} negative</span>
          </div>
        </div>"""

    # Open tasks table (top 10 by urgency)
    open_tasks = sorted(
        metrics.get("open_tasks", []),
        key=lambda x: (x.get("score", 5), -x.get("days_open", 0))
    )[:10]

    task_rows = ""
    region_badge_colors = {"South": "#2E7D32", "North": "#0D7C7C", "West": "#E65100"}
    for t in open_tasks:
        score = t.get("score", 0)
        days = t.get("days_open", 0)
        row_class = "row-critical" if score < 1.5 or days >= 14 else "row-warn"
        text = (t.get("text") or "")[:80] + ("..." if len(t.get("text") or "") > 80 else "")
        region = t.get("region", "")
        badge_color = region_badge_colors.get(region, "#455A64")
        day_color = "#B71C1C" if days >= 14 else "#E65100"
        score_color = rate_color(int(100 - score * 20))
        task_rows += f"""
        <tr class="{row_class}">
          <td>{t.get('location','')}</td>
          <td><span class="badge" style="background:{badge_color}">{region or '?'}</span></td>
          <td>{t.get('source','')}</td>
          <td style="color:{score_color}; font-weight:bold;">{score:.1f}</td>
          <td style="color:{day_color}; font-weight:bold;">{days}d</td>
          <td>{t.get('owner','Unassigned')}</td>
          <td class="review-text">{text}</td>
        </tr>"""

    net_rate = net.get("engagement_rate", 0)
    net_rate_color = rate_color(net_rate)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>WMCHealth CTL Report – {month_label}</title>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #F4F8FB; color: #455A64; }}

    /* Header */
    .header {{ background: #0D2B55; color: white; padding: 24px 40px; display: flex; justify-content: space-between; align-items: center; }}
    .header h1 {{ font-size: 22px; font-weight: 700; }}
    .header .meta {{ font-size: 12px; color: #A0C4D8; text-align: right; }}
    .header .download-btn {{ background: #0D7C7C; color: white; padding: 10px 20px; border-radius: 6px;
                             text-decoration: none; font-weight: 600; font-size: 13px; margin-left: 20px;
                             white-space: nowrap; transition: background 0.2s; }}
    .header .download-btn:hover {{ background: #0a6060; }}

    /* Main layout */
    .container {{ max-width: 1200px; margin: 0 auto; padding: 28px 24px; }}

    /* KPI Cards */
    .kpi-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 28px; }}
    .kpi-card {{ background: white; border-radius: 8px; padding: 20px; box-shadow: 0 2px 6px rgba(0,0,0,0.07);
                 border-top: 4px solid #ddd; }}
    .kpi-val {{ font-size: 36px; font-weight: 800; margin: 8px 0 4px; }}
    .kpi-label {{ font-size: 11px; font-weight: 700; color: #607D8B; text-transform: uppercase; letter-spacing: 0.5px; }}
    .kpi-sub {{ font-size: 11px; color: #90A4AE; margin-top: 4px; }}

    /* Section headings */
    .section-title {{ font-size: 16px; font-weight: 700; color: #0D2B55; margin: 28px 0 14px;
                      padding-bottom: 8px; border-bottom: 2px solid #E0E0E0; }}

    /* Region cards */
    .region-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 28px; }}
    .region-card {{ background: white; border-radius: 8px; padding: 18px; box-shadow: 0 2px 6px rgba(0,0,0,0.07); }}
    .region-name {{ font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }}
    .region-rate {{ font-size: 38px; font-weight: 800; margin: 8px 0 4px; }}
    .region-sub {{ font-size: 11px; color: #90A4AE; margin-bottom: 12px; }}
    .sentiment-bar {{ display: flex; height: 8px; border-radius: 4px; overflow: hidden; margin-bottom: 8px; }}
    .bar-pos {{ background: #2E7D32; }}
    .bar-neu {{ background: #E65100; }}
    .bar-neg {{ background: #B71C1C; }}
    .sentiment-legend {{ display: flex; gap: 12px; font-size: 10px; }}

    /* Open tasks table */
    .table-wrap {{ background: white; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.07); overflow: hidden; margin-bottom: 28px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
    thead tr {{ background: #0D2B55; color: white; }}
    th, td {{ padding: 10px 12px; text-align: left; border-bottom: 1px solid #ECEFF1; }}
    .row-critical {{ background: #FFEBEE; }}
    .row-warn {{ background: #FFF3E0; }}
    .review-text {{ color: #78909C; font-style: italic; max-width: 260px; }}
    .badge {{ color: white; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 700; }}

    /* Archive */
    .archive {{ background: white; border-radius: 8px; padding: 20px; box-shadow: 0 2px 6px rgba(0,0,0,0.07); }}
    .archive ul {{ list-style: none; padding: 0; }}
    .archive li {{ padding: 8px 0; border-bottom: 1px solid #ECEFF1; }}
    .archive a {{ color: #0D7C7C; text-decoration: none; font-weight: 500; }}
    .archive a:hover {{ text-decoration: underline; }}

    /* Footer */
    .footer {{ text-align: center; padding: 24px; font-size: 11px; color: #90A4AE; margin-top: 20px; }}

    @media (max-width: 768px) {{
      .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }}
      .region-grid {{ grid-template-columns: 1fr; }}
      .header {{ flex-direction: column; gap: 14px; text-align: center; }}
    }}
  </style>
</head>
<body>

<div class="header">
  <div>
    <div style="font-size:11px; color:#7EC8E3; letter-spacing:2px; margin-bottom:4px;">CLOSING THE LOOP</div>
    <h1>Engagement Report — {month_label}</h1>
    <div style="font-size:12px; color:#A0C4D8; margin-top:4px;">Westchester Medical Center Health Network</div>
  </div>
  <div style="display:flex; align-items:center; gap:16px;">
    <div class="meta">
      Source: Press Ganey<br/>
      Generated: {generated}
    </div>
    <a href="reports/{report_filename}" download class="download-btn">⬇ Download Report (.pptx)</a>
  </div>
</div>

<div class="container">

  <!-- KPI Cards -->
  <div class="kpi-grid">
    <div class="kpi-card" style="border-top-color:#0D7C7C;">
      <div class="kpi-label">Avg Engagement Rate</div>
      <div class="kpi-val" style="color:{net_rate_color};">{net_rate}%</div>
      <div class="kpi-sub">Network-wide</div>
    </div>
    <div class="kpi-card" style="border-top-color:#0D2B55;">
      <div class="kpi-label">Total Tasks</div>
      <div class="kpi-val" style="color:#0D2B55;">{net.get('total_tasks',0)}</div>
      <div class="kpi-sub">Total feedback records</div>
    </div>
    <div class="kpi-card" style="border-top-color:#2E7D32;">
      <div class="kpi-label">Closed Tasks</div>
      <div class="kpi-val" style="color:#2E7D32;">{net.get('closed_tasks',0)}</div>
      <div class="kpi-sub">Online / Email engaged</div>
    </div>
    <div class="kpi-card" style="border-top-color:#E65100;">
      <div class="kpi-label">Avg Time to Close</div>
      <div class="kpi-val" style="color:#E65100;">{fmt_hours(net.get('avg_time_to_close_hours',0))}</div>
      <div class="kpi-sub">From review to response</div>
    </div>
  </div>

  <!-- Regional Breakdown -->
  <div class="section-title">Regional Performance</div>
  <div class="region-grid">{region_cards}</div>

  <!-- Open / Critical Tasks -->
  <div class="section-title">⚠ Open & Critical Tasks (Top {len(open_tasks)} by Urgency)</div>
  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Location</th><th>Region</th><th>Source</th>
          <th>PFS Score</th><th>Days Open</th><th>Owner</th><th>Review Excerpt</th>
        </tr>
      </thead>
      <tbody>{task_rows}</tbody>
    </table>
  </div>

  <!-- Report Archive -->
  <div class="section-title">📁 Report Archive</div>
  <div class="archive">
    <ul>{archive_links or '<li>No previous reports yet.</li>'}</ul>
  </div>

</div>

<div class="footer">
  WMCHealth Closing the Loop Dashboard &nbsp;·&nbsp; Powered by Press Ganey API &nbsp;·&nbsp;
  Auto-generated by GitHub Actions &nbsp;·&nbsp; {generated}
</div>

</body>
</html>"""
    return html


def main():
    parser = argparse.ArgumentParser(description="Update GitHub Pages dashboard")
    parser.add_argument("--data", default="scripts/data.json")
    args = parser.parse_args()

    with open(args.data, "r") as f:
        metrics = json.load(f)

    month_str = metrics["meta"]["from_date"][:7]
    report_filename = f"WMCHealth_CTL_{month_str}.pptx"

    html = build_html(metrics, report_filename)

    os.makedirs("docs", exist_ok=True)
    with open("docs/index.html", "w") as f:
        f.write(html)

    print(f"✅ Dashboard updated: docs/index.html")
    print(f"   Report link: reports/{report_filename}")


if __name__ == "__main__":
    main()
