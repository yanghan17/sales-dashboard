"""
Monthly Drill-Down Dashboard
Run: python monthly_dashboard.py
Opens at http://localhost:5051
"""

import json
import os
import webbrowser
import threading
import time
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse
import pandas as pd
from datetime import datetime

# ─── CONFIG ────────────────────────────────────────────────────────────────────
_HERE     = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(_HERE, "sales_raw_sample.xlsx")
PORT      = 5051
VIEW_YEAR = None   # ← set to a specific year e.g. 2025, or None for latest year
# ───────────────────────────────────────────────────────────────────────────────


def load_data(filepath: str) -> pd.DataFrame:
    ext = os.path.splitext(filepath)[1].lower()
    df  = pd.read_csv(filepath) if ext == ".csv" else pd.read_excel(filepath)
    df.columns = df.columns.str.strip()
    df["Doc Date"] = pd.to_datetime(df["Doc Date"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["Doc Date"])
    df["Year"]      = df["Doc Date"].dt.year
    df["Month"]     = df["Doc Date"].dt.month
    df["Day"]       = df["Doc Date"].dt.day
    df["Weekday"]   = df["Doc Date"].dt.strftime("%a")
    df["MonthName"] = df["Doc Date"].dt.strftime("%B")
    df["DateStr"]   = df["Doc Date"].dt.strftime("%d %b")
    for col in ["Qty", "Unit Price", "Total"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def compute_year_data(df: pd.DataFrame, year: int, prev_year: int) -> dict:
    cy        = df[df["Year"] == year]
    py        = df[df["Year"] == prev_year] if prev_year in df["Year"].values else pd.DataFrame()
    agent_col = "Agent" if "Agent" in df.columns else None

    def safe_pct(a, b):
        return round((a - b) / b * 100, 1) if b else 0

    months_data = {}
    for m in range(1, 13):
        mdf       = cy[cy["Month"] == m]
        pmdf      = py[py["Month"] == m] if not py.empty else pd.DataFrame()
        month_name = datetime(year, m, 1).strftime("%B")

        if mdf.empty:
            months_data[m] = {"name": month_name, "has_data": False, "kpis": {},
                               "daily": [], "salesman": [], "top_customers": [],
                               "top_items": [], "transactions": [], "sm_daily": {}}
            continue

        revenue    = round(float(mdf["Total"].sum()), 2)
        py_revenue = round(float(pmdf["Total"].sum()), 2) if not pmdf.empty else 0
        orders     = int(len(mdf))
        customers  = int(mdf["Debtor Code"].nunique()) if "Debtor Code" in mdf.columns else 0
        avg_order  = round(revenue / orders, 2) if orders else 0
        growth     = safe_pct(revenue, py_revenue)
        bdg        = mdf.groupby("Day")["Total"].sum()
        best_day_val   = round(float(bdg.max()), 2) if not bdg.empty else 0
        best_day_num   = int(bdg.idxmax()) if not bdg.empty else 0
        best_day_label = datetime(year, m, best_day_num).strftime("%d %b") if best_day_num else "—"

        daily_grp = mdf.groupby(["Day", "DateStr", "Weekday"]).agg(
            Revenue=("Total", "sum"),
            Orders=("Doc Date", "count"),
            Customers=("Debtor Code", "nunique") if "Debtor Code" in mdf.columns else ("Total", "count"),
        ).reset_index().sort_values("Day")
        py_daily = pmdf.groupby("Day")["Total"].sum().to_dict() if not pmdf.empty else {}

        daily, cumulative = [], 0
        for _, row in daily_grp.iterrows():
            d   = int(row["Day"])
            rev = round(float(row["Revenue"]), 2)
            cumulative += rev
            daily.append({"day": d, "label": row["DateStr"], "weekday": row["Weekday"],
                           "revenue": rev, "orders": int(row["Orders"]), "customers": int(row["Customers"]),
                           "cumulative": round(cumulative, 2), "py_revenue": round(float(py_daily.get(d, 0)), 2)})

        salesman = []
        if agent_col:
            sm = mdf.groupby(agent_col).agg(Revenue=("Total","sum"), Orders=("Doc Date","count"),
                                             Customers=("Debtor Code","nunique")).reset_index().sort_values("Revenue", ascending=False)
            sm["AvgOrder"]    = (sm["Revenue"] / sm["Orders"]).round(2)
            sm["Revenue"]     = sm["Revenue"].round(2)
            sm["PrevRevenue"] = sm[agent_col].map(pmdf.groupby(agent_col)["Total"].sum()).fillna(0).round(2) if not pmdf.empty else 0
            sm["Growth"]      = sm.apply(lambda r: safe_pct(r["Revenue"], r["PrevRevenue"]), axis=1)
            salesman          = sm.rename(columns={agent_col: "Agent"}).to_dict("records")

        top_customers = []
        if "Debtor Name" in mdf.columns:
            cust = mdf.groupby("Debtor Name").agg(Revenue=("Total","sum"), Orders=("Doc Date","count")) \
                      .reset_index().sort_values("Revenue", ascending=False).head(10)
            cust["Revenue"] = cust["Revenue"].round(2)
            top_customers   = cust.to_dict("records")

        top_items = []
        if "Item Code" in mdf.columns:
            items = mdf.groupby("Item Code").agg(Revenue=("Total","sum"), Qty=("Qty","sum"), Orders=("Doc Date","count")) \
                       .reset_index().sort_values("Revenue", ascending=False).head(10)
            items["Revenue"] = items["Revenue"].round(2)
            top_items        = items.rename(columns={"Item Code": "Item"}).to_dict("records")

        tx_cols      = [c for c in ["Doc Date","Debtor Name","Agent","Item Code","Detail Description","Qty","Unit Price","Total"] if c in mdf.columns]
        tx           = mdf[tx_cols].copy()
        tx["Doc Date"] = tx["Doc Date"].dt.strftime("%d %b %Y")
        transactions = tx.sort_values("Doc Date", ascending=False).head(200).to_dict("records")

        sm_daily = {}
        if agent_col:
            day_nums = [r["day"] for r in daily]
            agents   = [s["Agent"] for s in salesman]
            sm_daily = {"agents": agents, "days": [r["label"] for r in daily], "series": []}
            for ag in agents:
                ag_day = mdf[mdf[agent_col] == ag].groupby("Day")["Total"].sum()
                sm_daily["series"].append({"name": ag, "data": [round(float(ag_day.get(d, 0)), 2) for d in day_nums]})

        months_data[m] = {
            "name": month_name, "has_data": True,
            "kpis": {"revenue": revenue, "py_revenue": py_revenue, "growth": growth,
                     "orders": orders, "customers": customers, "avg_order": avg_order,
                     "best_day_val": best_day_val, "best_day_label": best_day_label},
            "daily": daily, "sm_daily": sm_daily, "salesman": salesman,
            "top_customers": top_customers, "top_items": top_items, "transactions": transactions,
        }

    month_overview = [{"month": m, "name": datetime(year, m, 1).strftime("%b"),
                        "revenue": round(float(cy[cy["Month"]==m]["Total"].sum()), 2),
                        "orders":  int(len(cy[cy["Month"]==m]))} for m in range(1, 13)]

    return {"months": {str(k): v for k, v in months_data.items()}, "month_overview": month_overview}


def build_initial_payload(df: pd.DataFrame) -> dict:
    all_years    = sorted(df["Year"].unique().tolist())
    current_year = int(VIEW_YEAR) if VIEW_YEAR and int(VIEW_YEAR) in all_years else max(all_years)
    data         = compute_year_data(df, current_year, current_year - 1)
    return {
        "meta": {
            "current_year": current_year,
            "all_years":    all_years,
            "generated":    datetime.now().strftime("%d %b %Y %H:%M"),
            "file":         os.path.basename(DATA_FILE),
        },
        "month_overview": data["month_overview"],
        "months":         data["months"],
    }


def get_html(data: dict) -> str:
    data_json = json.dumps(data, ensure_ascii=False)
    with open(os.path.join(_HERE, "monthly_template.html"), "r", encoding="utf-8") as f:
        html = f.read()
    return html.replace("__DATA_PLACEHOLDER__", data_json)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args): pass

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path in ("/", "/index.html"):
            try:
                df   = load_data(DATA_FILE)
                data = build_initial_payload(df)
                html = get_html(data)
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(html.encode("utf-8"))
            except FileNotFoundError:
                self.send_response(404); self.send_header("Content-Type","text/html"); self.end_headers()
                self.wfile.write(f"<pre>File not found: {DATA_FILE}</pre>".encode())
            except Exception as e:
                self.send_response(500); self.send_header("Content-Type","text/html"); self.end_headers()
                self.wfile.write(f"<pre>Error: {e}</pre>".encode())
            return
        self.send_response(404); self.end_headers()


def open_browser():
    time.sleep(0.8)
    webbrowser.open(f"http://localhost:{PORT}")


if __name__ == "__main__":
    print(f"\n{'─'*50}")
    print(f"  📅 Monthly Drill-Down Dashboard")
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