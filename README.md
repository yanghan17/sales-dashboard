# 📊 Local Sales Analytics Dashboard

A self-hosted business intelligence dashboard built in Python — no cloud, no subscriptions, no data leaving your machine. Designed for Malaysian SME distributors and trading companies running AutoCount or SQL Account.

Built as a commercial product targeting businesses that export invoice data to Excel but have no real-time visibility into their sales performance.

![Python](https://img.shields.io/badge/Python-3.8+-blue?style=flat-square&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey?style=flat-square)

---

## The Problem

Most SME distributors in Malaysia use accounting software (AutoCount, SQL Account) that stores years of transaction data — but provides little to no meaningful sales analytics. Business owners rely on month-end reports, verbal updates from salesmen, or manually scrolling through Excel to understand their business.

This dashboard changes that. Point it at an Excel export, launch it, and get a full picture of revenue trends, salesman performance, customer behaviour, and product movement — instantly, locally, privately.

---

## Features

### Main Dashboard (`dashboard.py` → `localhost:5050`)
- **KPI Cards** — Total revenue, YoY growth, invoice count, average order value, unique customers
- **Monthly Revenue Chart** — Current vs previous year side by side, with year tab switching
- **Yearly Revenue Bar Chart** — Full historical trend across all years in data
- **Salesman Leaderboard** — Revenue, YoY growth, orders, customer count, market share bar per agent
- **Salesman Monthly Chart** — Line chart showing each agent's monthly contribution over time
- **Top 10 Customers** — Ranked by revenue for selected year
- **Top 10 Products** — Ranked by revenue for selected year
- **Revenue Heatmap** — All years × 12 months colour intensity grid, spot seasonal patterns instantly
- **Monthly Summary Table** — Full history with revenue, orders, unique customers per month

### Monthly Drill-Down (`monthly_dashboard.py` → `localhost:5051`)
- **12-Month Overview Cards** — Visual revenue bars showing relative performance, click to drill in
- **Daily Revenue Chart** — Day-by-day breakdown vs same month last year
- **Cumulative Revenue Chart** — Running total through the month
- **Daily Sales by Agent** — Stacked bar chart per day per salesman
- **Agent Performance Table** — Revenue, YoY growth, orders, customers, avg order, market share
- **Top Customers & Products** — For that specific month
- **Daily Summary Table** — Every trading day with DoD comparison
- **Transaction Log** — Latest 200 rows for full auditability

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3 standard library (`http.server`, `json`, `threading`) |
| Data processing | `pandas` |
| File support | `openpyxl` (xlsx), `xlrd` (xls) |
| Frontend | Vanilla HTML + CSS + JavaScript |
| Charts | Chart.js 4.4 (CDN) |
| Fonts | Inter + IBM Plex Mono (Google Fonts) |

Zero framework dependencies. No Flask, no Django, no Node. The entire backend is Python stdlib — intentional, for dead-simple deployment on client machines.

---

## Quick Start

**1. Install dependencies**
```bash
pip install pandas openpyxl xlrd
```

**2. Add your data file**

Put your Excel export in the same folder as the scripts. Open `dashboard.py` and set:
```python
DATA_FILE = os.path.join(_HERE, "your_file.xlsx")
```

**3. Launch**
```bash
python dashboard.py          # Main dashboard  → http://localhost:5050
python monthly_dashboard.py  # Monthly drill-down → http://localhost:5051
```
Browser opens automatically.

**4. Try with sample data**
```bash
python generate_sample_data.py   # generates sales_raw_sample.xlsx
python dashboard.py
```

---

## Data Format

Your Excel file must have these columns (AutoCount and SQL Account export these by default):

| Column | Format | Notes |
|---|---|---|
| `Doc Date` | DD/MM/YYYY | Invoice date |
| `Debtor Code` | Text | Customer ID |
| `Debtor Name` | Text | Customer name |
| `Agent` | Text | Salesman name |
| `Item Code` | Text | SKU / product code |
| `Detail Description` | Text | Product name |
| `Qty` | Number | Quantity |
| `Unit Price` | Number | Price per unit |
| `Total` | Number | Line total (negative = return/credit note) |

Column names are case-sensitive. Extra whitespace in headers is stripped automatically.

---

## Architecture

```
Excel File (.xlsx / .csv)
        │
        ▼
  pandas DataFrame
        │
        ▼
  compute engine (Python)
  ├── KPIs
  ├── Monthly aggregations
  ├── Salesman breakdowns
  ├── Top customers / products
  ├── Daily series
  └── Heatmap data
        │
        ▼
  JSON payload (injected into HTML template)
        │
        ▼
  Python HTTP server (stdlib)
        │
        ▼
  Browser (Chart.js renders everything client-side)
```

All computation happens in Python before the page loads. The browser receives pre-computed JSON and renders it — no API calls, no database, no backend logic in JavaScript.

---

## Configuration

```python
# dashboard.py
DATA_FILE = os.path.join(_HERE, "sales_raw.xlsx")   # ← your data file
PORT      = 5050                                      # ← change if port in use

# monthly_dashboard.py
DATA_FILE = os.path.join(_HERE, "sales_raw.xlsx")
PORT      = 5051
VIEW_YEAR = None   # ← set to e.g. 2024 to pin a specific year, None = latest
```

---

## Sample Data

`generate_sample_data.py` creates a realistic dataset representing a Malaysian metalworking consumables distributor:

- 5,100+ rows across 2024 (full year) and 2025 (Jan–Oct)
- 4 sales agents with territory-based customer affinity
- 15 customers (Malaysian SME-style company names)
- 15 SKUs — end mills, carbide drills, CNMG inserts, grinding wheels, cutting oil
- Seasonal patterns, ~4% returns, 12% YoY growth baked in

---

## Roadmap

This dashboard is Plan A of a three-tier commercial product:

- **Plan A (this repo)** — Local dashboard, revenue and salesman visibility
- **Plan B** — Customer risk flags, last order date tracking, SKU-level trends, churn early warning
- **Plan C** — AI integration via Claude API, proactive Telegram alerts, natural language queries

Plan C builds on [BizBot](https://github.com/yanghan17/bizbot) — a separate repo implementing the AI and Telegram layer.

---

## Why Local-Only

Privacy is a first-class feature. Malaysian SME owners are uncomfortable putting customer names, revenue figures, and business performance data on foreign cloud servers. Everything runs on the client's own machine — no accounts, no subscriptions, no data in transit.

---

## License

MIT
