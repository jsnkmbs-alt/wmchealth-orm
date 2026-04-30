# WMCHealth ORM — CTL Report Pipeline

Automatically generates the **Closing the Loop Engagement Report** PowerPoint
from live Press Ganey API data, hosted as a GitHub Pages dashboard.

---

## How It Works

```
Press Ganey API
      ↓  (1st of every month, 8am UTC — or manual trigger)
GitHub Actions:
   1. fetch_data.py       → pulls reviews + engagement → saves data.json
   2. build_report.py     → reads data.json → generates .pptx
   3. update_dashboard.py → reads data.json → updates docs/index.html
      ↓
GitHub Pages:
   https://YOUR-USERNAME.github.io/wmchealth-orm/
```

---

## First-Time Setup

### 1. Add Secrets
Go to **Settings → Secrets and variables → Actions → New repository secret**

| Secret Name | Value |
|---|---|
| `PG_APP_ID` | Your Press Ganey Application ID |
| `PG_APP_SECRET` | Your Press Ganey Application Secret |
| `PG_NODE_ID` | Your Press Ganey Node ID (optional) |

### 2. Enable GitHub Pages
**Settings → Pages → Source: Deploy from branch → main / /orm-reports/docs → Save**

### 3. Run Your First Report
**Actions → Monthly CTL Report → Run workflow → enter month (e.g. `2026-03`) → Run**

### 4. Update Region Map (after first run)
Open `orm-reports/scripts/region_map.json` and add real location IDs
from `reviews[].location.id` in the API response to the `location_id_map` section.

---

## File Structure

```
wmchealth-orm/                          ← GitHub repository root
├── .github/
│   └── workflows/
│       └── monthly_report.yml          ← Scheduler & pipeline
└── orm-reports/
    ├── requirements.txt                ← Python dependencies
    ├── scripts/
    │   ├── fetch_data.py               ← Press Ganey API → data.json
    │   ├── build_report.py             ← data.json → .pptx
    │   ├── update_dashboard.py         ← data.json → index.html
    │   ├── region_map.json             ← Location → Region mapping
    │   └── data.json                   ← Auto-generated (do not edit)
    └── docs/                           ← GitHub Pages root
        ├── index.html                  ← Auto-generated dashboard
        └── reports/
            └── WMCHealth_CTL_YYYY-MM.pptx
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Workflow not in Actions tab | Check `.github/workflows/monthly_report.yml` exists |
| `PG_APP_ID … required` error | Add secrets in Settings → Secrets → Actions |
| Auth failed | Double-check credentials with Press Ganey Account Manager |
| Regions show "Unknown" | Update `region_map.json` with real location IDs |
| GitHub Pages not live | Settings → Pages → Source: main / `/orm-reports/docs` |
