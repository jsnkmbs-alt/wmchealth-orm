# WMCHealth CTL Report — Automated Pipeline

Automatically generates the **Closing the Loop Engagement Report** PowerPoint deck
from live Press Ganey API data, hosted as a GitHub Pages dashboard.

---

## How It Works

```
Press Ganey API
      ↓  (1st of every month, 8am)
GitHub Actions runs:
   1. fetch_data.py      → pulls reviews + engagement data → saves data.json
   2. build_report.py    → reads data.json → generates .pptx
   3. update_dashboard.py → reads data.json → updates docs/index.html
      ↓
GitHub Pages
   → Live dashboard at https://<your-org>.github.io/wmchealth-ctl-reports/
   → PPTX download available on the dashboard
```

---

## First-Time Setup (Do This Once)

### Step 1 — Create the GitHub Repository

1. Go to [github.com](https://github.com) → click **"New repository"**
2. Name it: `wmchealth-ctl-reports`
3. Set visibility to **Public**
4. Click **"Create repository"**

### Step 2 — Upload These Files

Option A — GitHub web interface (easiest):
1. Click **"uploading an existing file"** on the empty repo page
2. Drag and drop this entire folder
3. Click **"Commit changes"**

Option B — Git command line:
```bash
git clone https://github.com/YOUR-ORG/wmchealth-ctl-reports.git
cd wmchealth-ctl-reports
# copy these files in, then:
git add .
git commit -m "Initial setup"
git push
```

### Step 3 — Add Your Press Ganey Secrets

1. Go to your repo on GitHub
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Click **"New repository secret"** and add each of these:

| Secret Name | Value |
|---|---|
| `PG_APP_ID` | Your Press Ganey Application ID |
| `PG_APP_SECRET` | Your Press Ganey Application Secret |
| `PG_NODE_ID` | Your Press Ganey node ID (optional — for engagement rate API) |

### Step 4 — Enable GitHub Pages

1. Go to **Settings** → **Pages**
2. Under **Source**, select **"Deploy from a branch"**
3. Branch: `main` / Folder: `/docs`
4. Click **Save**

Your dashboard will be live at:
`https://<your-github-username>.github.io/wmchealth-ctl-reports/`

### Step 5 — Update the Region Map

After your first API call (Step 6), you'll have real location IDs.
Open `scripts/region_map.json` and populate the `location_id_map` section
with the IDs returned in `reviews[].location.id` from the API.

### Step 6 — Run Your First Report Manually

1. Go to the **Actions** tab in your repo
2. Click **"Monthly CTL Report"** in the left sidebar
3. Click **"Run workflow"**
4. Optionally enter a month (e.g. `2026-03`) or leave blank for last month
5. Click **"Run workflow"** (green button)

Watch the logs — within ~2 minutes your dashboard and PPTX will be live!

---

## Ongoing — Fully Automatic

After setup, the pipeline runs automatically on the **1st of every month at 8am UTC**.
No manual action required. Each run:
- Pulls fresh data from Press Ganey
- Generates a new `.pptx` report
- Updates the dashboard
- Archives the previous report

---

## Manual Run for a Specific Month

Go to **Actions** → **Monthly CTL Report** → **Run workflow** → enter `2026-03` (for example).

---

## File Structure

```
wmchealth-ctl-reports/
│
├── .github/workflows/
│   └── monthly_report.yml      ← Scheduler (runs on 1st of month)
│
├── scripts/
│   ├── fetch_data.py           ← Calls Press Ganey API, calculates metrics
│   ├── build_report.py         ← Generates PowerPoint from metrics
│   ├── update_dashboard.py     ← Generates GitHub Pages HTML dashboard
│   ├── region_map.json         ← Maps location names/IDs to regions
│   └── data.json               ← Auto-generated metrics (do not edit manually)
│
├── docs/                       ← GitHub Pages root
│   ├── index.html              ← Auto-generated dashboard
│   └── reports/
│       └── WMCHealth_CTL_YYYY-MM.pptx  ← Auto-generated reports
│
├── requirements.txt            ← Python dependencies
└── README.md                   ← This file
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `PG_APP_ID and PG_APP_SECRET environment variables are required` | Add secrets in Settings → Secrets → Actions |
| `Auth failed` | Double-check credentials with your Press Ganey Account Manager |
| Report shows 0 tasks | Check date range — make sure reviews exist for that month |
| GitHub Pages not showing | Wait 2-3 minutes after first push; check Settings → Pages |
| Regions all show "Unknown" | Update `region_map.json` with real location IDs from your first API run |

---

## Dependencies

- `requests` — HTTP calls to Press Ganey API
- `python-pptx` — PowerPoint generation
- `python-dateutil` — Date calculations
- `Pillow` — Image processing support for python-pptx
