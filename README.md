# AIDA Ventures — Integrated Intelligence Dashboard

A unified static dashboard combining fintech market intelligence, VC fund benchmarks, deal evaluation tools and venture studio metrics for AIDA Ventures' hybrid VC + Studio model across LatAm.

## Live Dashboard

Deployable on GitHub Pages → `outputs/index.html`

---

## Project Structure

```
aida_integrated_dashboard/
│
├── data/
│   ├── raw/                       ← Excel source files (5 workbooks)
│   └── processed/
│       └── dashboard_data.json    ← Auto-generated from raw files
│
├── inputs/                        ← Original HTML dashboards (reference)
│
├── src/
│   ├── build_dashboard.py         ← Main builder — generates outputs/index.html
│   ├── extract_data.py            ← Reads all Excel files → structured dict
│   └── utils.py                   ← Helper functions (clean values, safe str)
│
├── outputs/
│   └── index.html                 ← Final dashboard (generated, deploy this)
│
├── README.md
├── requirements.txt
└── .gitignore
```

---

## Dashboard Sections

| # | Section | Source |
|---|---------|--------|
| 1 | **Overview** — Hero KPIs + charts | All sources |
| 2 | **Market Intelligence** — Fintech sectors LATAM & US | `Fintech_Sectors.xlsx` |
| 3 | **VC Benchmarks** — IRR, TVPI, DPI, valuations, timing | `VCFunds_Metrics.xlsx` + `_AIDA_Ventures_-_Startups_Benchmarks.xlsx` |
| 4 | **Deal Scorecard** — Interactive 0–100 scoring | `_Metricas_Startups.xlsx` |
| 5 | **Venture Studio Metrics** — A/B/C metrics reference | `Venture_Studio_Metrics_Reference.xlsx` |
| 6 | **Methodology** — Data sources + generation date | — |

---

## Setup

### 1. Clone the repo

```bash
git clone https://github.com/YOUR_USERNAME/aida-integrated-dashboard.git
cd aida-integrated-dashboard
```

### 2. Create virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate       # macOS/Linux
.venv\Scripts\activate          # Windows
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Place Excel files

Ensure these files are in `data/raw/`:
- `Fintech_Sectors.xlsx`
- `VCFunds_Metrics.xlsx`
- `_Metricas_Startups.xlsx`
- `Venture_Studio_Metrics_Reference.xlsx`
- `_AIDA_Ventures_-_Startups_Benchmarks.xlsx`

### 5. Build the dashboard

```bash
cd src
python3 build_dashboard.py
```

This generates `outputs/index.html`.

### 6. Open locally

```bash
open ../outputs/index.html       # macOS
start ../outputs/index.html      # Windows
xdg-open ../outputs/index.html   # Linux
```

---

## GitHub Pages Deployment

### Option A — Automated (recommended)

1. Push `outputs/index.html` to `main` branch
2. Go to repo **Settings → Pages**
3. Source: **Deploy from a branch** → `main` → `/outputs` folder
4. Save → your dashboard is live at `https://YOUR_USERNAME.github.io/aida-integrated-dashboard/`

### Option B — Copy to root

```bash
cp outputs/index.html ./index.html
git add index.html
git commit -m "deploy: update dashboard"
git push
```

Then set GitHub Pages source to `main` branch, root folder.

---

## Git Workflow

### Suggested branch structure

```
main          ← stable / production (what GitHub Pages serves)
dev           ← integration branch
feature/...   ← feature development
```

### Suggested commits

```bash
git init
git add .
git commit -m "init: project structure"

git add data/raw/
git commit -m "add: raw excel sources and original HTML inputs"

git add src/extract_data.py src/utils.py
git commit -m "add: data extraction pipeline"

git add src/build_dashboard.py
git commit -m "add: integrated dashboard builder"

python3 src/build_dashboard.py
git add outputs/index.html data/processed/
git commit -m "generate: first static index.html"

git push -u origin main
```

---

## Updating the Dashboard

When source Excel files change:

```bash
# 1. Replace files in data/raw/
# 2. Rebuild
cd src && python3 build_dashboard.py
# 3. Commit and push
git add outputs/index.html data/processed/dashboard_data.json
git commit -m "update: regenerate dashboard from new data"
git push
```

---

## Technical Notes

- **Single-file output** — `index.html` is self-contained (CSS + JS embedded)
- **CDN dependencies** — Chart.js 4.4.1, Google Fonts (requires internet)
- **No backend** — pure static HTML, zero server requirements
- **Chart.js charts** — 12 interactive charts across all sections
- **Interactive scorecard** — 9-input real-time scoring engine
- **Responsive** — mobile-friendly layout with breakpoint at 900px
- **GitHub Pages compatible** — no build step, no absolute paths

---

*AIDA Ventures — Intelligence Platform · Static dashboard · GitHub Pages ready*
