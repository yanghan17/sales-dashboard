"""
Invoice Dashboard - Live Analytics
Run: python dashboard.py
Opens a live dashboard in your browser at http://localhost:5050
"""

import json
import os
import webbrowser
import threading
import time
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import pandas as pd
import numpy as np
from datetime import datetime

# ─── CONFIG ────────────────────────────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(_HERE, "sales_raw_sample.xlsx")   # ← change filename here
PORT = 5050
# ───────────────────────────────────────────────────────────────────────────────


def load_data(filepath: str) -> pd.DataFrame:
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(filepath)
    else:
        df = pd.read_excel(filepath)

    # Normalise column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Parse Doc Date
    df["Doc Date"] = pd.to_datetime(df["Doc Date"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["Doc Date"])

    df["Year"]  = df["Doc Date"].dt.year
    df["Month"] = df["Doc Date"].dt.month
    df["MonthName"] = df["Doc Date"].dt.strftime("%b")
    df["YearMonth"] = df["Doc Date"].dt.strftime("%Y-%m")
    df["YearMonthLabel"] = df["Doc Date"].dt.strftime("%b %Y")

    # Coerce numeric
    for col in ["Qty", "Unit Price", "Total"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def compute_analytics(df: pd.DataFrame) -> dict:
    years = sorted(df["Year"].unique().tolist())
    current_year = max(years)
    cy = df[df["Year"] == current_year]
    py = df[df["Year"] == current_year - 1] if (current_year - 1) in years else pd.DataFrame()

    def safe_pct(a, b):
        return round((a - b) / b * 100, 1) if b else 0

    # ── KPI Cards ──────────────────────────────────────────────────────────────
    total_revenue   = round(float(df["Total"].sum()), 2)
    cy_revenue      = round(float(cy["Total"].sum()), 2)
    py_revenue      = round(float(py["Total"].sum()), 2) if not py.empty else 0
    revenue_growth  = safe_pct(cy_revenue, py_revenue)

    total_invoices  = int(df["Doc Date"].count())
    cy_invoices     = int(cy["Doc Date"].count())

    avg_order_value = round(cy_revenue / cy_invoices, 2) if cy_invoices else 0

    unique_customers = int(df["Debtor Code"].nunique()) if "Debtor Code" in df.columns else 0

    # ── Monthly Revenue (current year) ────────────────────────────────────────
    monthly = (
        cy.groupby(["Month", "MonthName"])["Total"]
        .sum().reset_index()
        .sort_values("Month")
    )
    month_labels  = monthly["MonthName"].tolist()
    month_values  = [round(v, 2) for v in monthly["Total"].tolist()]

    # prev year same months
    py_monthly = {}
    if not py.empty:
        pm = py.groupby("Month")["Total"].sum()
        py_monthly = {int(k): round(float(v), 2) for k, v in pm.items()}
    py_month_values = [py_monthly.get(m, 0) for m in monthly["Month"].tolist()]

    # ── Yearly Revenue ─────────────────────────────────────────────────────────
    yearly = df.groupby("Year")["Total"].sum().reset_index().sort_values("Year")
    year_labels = [str(y) for y in yearly["Year"].tolist()]
    year_values = [round(v, 2) for v in yearly["Total"].tolist()]

    # ── Salesman Report (ALL years) ───────────────────────────────────────────
    agent_col = "Agent" if "Agent" in df.columns else None
    salesman_data = []
    salesman_all = {}
    if agent_col:
        for yr in years:
            ydf = df[df["Year"] == yr]
            prev_ydf = df[df["Year"] == yr - 1]
            sm = ydf.groupby(agent_col).agg(
                Revenue=("Total", "sum"),
                Orders=("Doc Date", "count"),
                Customers=("Debtor Code", "nunique")
            ).reset_index().sort_values("Revenue", ascending=False)
            sm["AvgOrder"] = (sm["Revenue"] / sm["Orders"]).round(2)
            sm["Revenue"] = sm["Revenue"].round(2)
            if not prev_ydf.empty:
                py_sm = prev_ydf.groupby(agent_col)["Total"].sum()
                sm["PrevRevenue"] = sm[agent_col].map(py_sm).fillna(0).round(2)
                sm["Growth"] = sm.apply(
                    lambda r: safe_pct(r["Revenue"], r["PrevRevenue"]), axis=1
                )
            else:
                sm["PrevRevenue"] = 0
                sm["Growth"] = 0
            salesman_all[str(yr)] = sm.rename(columns={agent_col: "Agent"}).to_dict("records")
        salesman_data = salesman_all.get(str(current_year), [])

    # ── Monthly Salesman Breakdown (ALL years) ───────────────────────────────
    sm_monthly = {}
    sm_monthly_all = {}
    agents_all = sorted(df[agent_col].dropna().unique().tolist()) if agent_col else []
    if agent_col:
        for yr in years:
            ydf = df[df["Year"] == yr]
            yr_monthly = (
                ydf.groupby(["Month", "MonthName"])["Total"]
                .sum().reset_index().sort_values("Month")
            )
            yr_month_labels = yr_monthly["MonthName"].tolist()
            yr_month_nums   = yr_monthly["Month"].tolist()
            smc = ydf.groupby([agent_col, "Month"])["Total"].sum().reset_index()
            series = []
            for ag in agents_all:
                ag_data = smc[smc[agent_col] == ag].set_index("Month")["Total"]
                vals = [round(float(ag_data.get(m, 0)), 2) for m in yr_month_nums]
                if any(v > 0 for v in vals):
                    series.append({"name": ag, "data": vals})
            sm_monthly_all[str(yr)] = {
                "agents": agents_all,
                "months": yr_month_labels,
                "series": series,
            }
        sm_monthly = sm_monthly_all.get(str(current_year), {})

    # ── Top Products (ALL years) ──────────────────────────────────────────────
    item_col = "Item Code" if "Item Code" in df.columns else None
    top_items = []
    top_items_all = {}
    if item_col:
        for yr in years:
            ydf = df[df["Year"] == yr]
            items = ydf.groupby(item_col).agg(
                Revenue=("Total", "sum"),
                Qty=("Qty", "sum"),
                Orders=("Doc Date", "count")
            ).reset_index().sort_values("Revenue", ascending=False).head(10)
            items["Revenue"] = items["Revenue"].round(2)
            top_items_all[str(yr)] = items.rename(columns={item_col: "Item"}).to_dict("records")
        top_items = top_items_all.get(str(current_year), [])

    # ── Top Customers (ALL years) ──────────────────────────────────────────────
    top_customers = []
    top_customers_all = {}
    if "Debtor Name" in df.columns:
        for yr in years:
            ydf = df[df["Year"] == yr]
            cust = ydf.groupby("Debtor Name").agg(
                Revenue=("Total", "sum"),
                Orders=("Doc Date", "count")
            ).reset_index().sort_values("Revenue", ascending=False).head(10)
            cust["Revenue"] = cust["Revenue"].round(2)
            top_customers_all[str(yr)] = cust.to_dict("records")
        top_customers = top_customers_all.get(str(current_year), [])

    # ── Monthly Transactions (for table) ──────────────────────────────────────
    monthly_summary = []
    for yr in years:
        ydf = df[df["Year"] == yr]
        for m in range(1, 13):
            mdf = ydf[ydf["Month"] == m]
            if len(mdf) == 0:
                continue
            label = datetime(yr, m, 1).strftime("%b %Y")
            monthly_summary.append({
                "label": label,
                "year": yr,
                "month": m,
                "revenue": round(float(mdf["Total"].sum()), 2),
                "orders": int(len(mdf)),
                "customers": int(mdf["Debtor Code"].nunique()) if "Debtor Code" in mdf.columns else 0,
            })

    # ── Revenue by Month Heatmap ───────────────────────────────────────────────
    heatmap = []
    for yr in years:
        row = {"year": yr, "months": []}
        ydf = df[df["Year"] == yr]
        for m in range(1, 13):
            val = ydf[ydf["Month"] == m]["Total"].sum()
            row["months"].append(round(float(val), 2))
        heatmap.append(row)

    return {
        "meta": {
            "years": years,
            "current_year": current_year,
            "generated": datetime.now().strftime("%d %b %Y %H:%M"),
            "file": os.path.basename(DATA_FILE),
        },
        "kpis": {
            "total_revenue": total_revenue,
            "cy_revenue": cy_revenue,
            "py_revenue": py_revenue,
            "revenue_growth": revenue_growth,
            "total_invoices": total_invoices,
            "cy_invoices": cy_invoices,
            "avg_order_value": avg_order_value,
            "unique_customers": unique_customers,
        },
        "monthly": {
            "labels": month_labels,
            "values": month_values,
            "py_values": py_month_values,
        },
        "yearly": {
            "labels": year_labels,
            "values": year_values,
        },
        "salesman": salesman_data,
        "salesman_all": salesman_all,
        "sm_monthly": sm_monthly,
        "sm_monthly_all": sm_monthly_all,
        "top_items": top_items,
        "top_items_all": top_items_all,
        "top_customers": top_customers,
        "top_customers_all": top_customers_all,
        "monthly_summary": monthly_summary,
        "heatmap": heatmap,
    }


# ── Read HTML template ─────────────────────────────────────────────────────────
def get_html(data: dict) -> str:
    data_json = json.dumps(data, ensure_ascii=False)
    with open(os.path.join(_HERE, "dashboard_template.html"), "r", encoding="utf-8") as f:
        html = f.read()
    return html.replace("__DATA_PLACEHOLDER__", data_json)


# ── HTTP Server ────────────────────────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # silence access logs

    def do_GET(self):
        parsed = urlparse(self.path)

        if parsed.path == "/data":
            # Live reload endpoint
            try:
                df = load_data(DATA_FILE)
                analytics = compute_analytics(df)
                payload = json.dumps(analytics).encode()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(payload)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())
            return

        if parsed.path == "/" or parsed.path == "/index.html":
            try:
                df = load_data(DATA_FILE)
                analytics = compute_analytics(df)
                html = get_html(analytics)
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(html.encode("utf-8"))
            except FileNotFoundError:
                page = f"""<html><body style="font-family:monospace;padding:40px;background:#0f0f0f;color:#ff4444">
                <h2>⚠ File not found: <code>{DATA_FILE}</code></h2>
                <p>Edit <code>dashboard.py</code> and set <code>DATA_FILE</code> to your Excel or CSV path.</p>
                <p>Expected columns: Doc Date, Debtor Code, Debtor Name, Agent, Item Code, Detail Description, Qty, Unit Price, Total</p>
                </body></html>"""
                self.send_response(404)
                self.send_header("Content-Type", "text/html")
                self.end_headers()
                self.wfile.write(page.encode())
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "text/html")
                self.end_headers()
                self.wfile.write(f"<pre>Error: {e}</pre>".encode())
            return

        self.send_response(404)
        self.end_headers()


def open_browser():
    time.sleep(0.8)
    webbrowser.open(f"http://localhost:{PORT}")


if __name__ == "__main__":
    print(f"\n{'─'*50}")
    print(f"  📊 Invoice Dashboard")
    print(f"  Data file : {DATA_FILE}")
    print(f"  URL       : http://localhost:{PORT}")
    print(f"  Stop      : Ctrl+C")
    print(f"{'─'*50}\n")
    threading.Thread(target=open_browser, daemon=True).start()
    server = HTTPServer(("localhost", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")