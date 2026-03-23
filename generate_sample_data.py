"""
generate_sample_data.py
Run this once to create a realistic invoice_data.xlsx for testing.
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

random.seed(42); np.random.seed(42)

AGENTS   = ["Ahmad Rizal", "Tan Wei Ming", "Priya Nair", "Lim Chee Keong", "Hafiz Sulaiman"]
DEBTORS  = {
    "D001":"Syarikat Teknik Jaya Sdn Bhd","D002":"CNC Parts Solutions Sdn Bhd",
    "D003":"Metal Works KL Sdn Bhd","D004":"Precision Engineering Sdn Bhd",
    "D005":"TechFab Industries Sdn Bhd","D006":"Keystone Machinery Sdn Bhd",
    "D007":"Borneo Steel Works Sdn Bhd","D008":"Delta Engineering Sdn Bhd",
    "D009":"Advance Machining Sdn Bhd","D010":"Prime Tooling Sdn Bhd",
    "D011":"Malaya Sheet Metal Sdn Bhd","D012":"KL Fabricators Sdn Bhd",
    "D013":"Shah Alam Metal Sdn Bhd","D014":"Penang Precision Sdn Bhd",
    "D015":"JB Engineering Supplies Sdn Bhd",
}
ITEMS = {
    "TL-001":("Carbide End Mill 10mm","pcs",45.00),
    "TL-002":("HSS Drill Bit Set 1-13mm","set",120.00),
    "TL-003":("Indexable Insert CNMG 120408","box",180.00),
    "TL-004":("Face Milling Cutter 80mm","pcs",380.00),
    "TL-005":("Boring Bar 32mm Shank","pcs",250.00),
    "MC-001":("Coolant Concentrate 20L","drum",160.00),
    "MC-002":("Machine Oil ISO 46 20L","drum",95.00),
    "MC-003":("Cutting Fluid EP 5L","can",48.00),
    "EQ-001":("Digital Vernier Caliper 0-150mm","pcs",75.00),
    "EQ-002":("Magnetic Base Indicator Stand","pcs",135.00),
    "EQ-003":("Dial Test Indicator 0.01mm","pcs",90.00),
    "SP-001":("CNC Chuck Jaw Set 3pcs","set",220.00),
    "SP-002":("Collet Chuck ER32 Set","set",310.00),
    "SP-003":("Quick Change Tool Holder BT40","pcs",185.00),
    "SP-004":("Linear Guide Rail 400mm","pcs",280.00),
}

rows = []
agent_debtor = {a: random.sample(list(DEBTORS.keys()), k=random.randint(5,10)) for a in AGENTS}

for year in range(2020, 2026):
    for month in range(1, 13):
        growth = 1.0 + (year - 2020)*0.08
        n_invoices = int(random.gauss(60, 12) * growth)
        n_invoices = max(20, n_invoices)
        for _ in range(n_invoices):
            agent = random.choices(AGENTS, weights=[0.25,0.22,0.18,0.20,0.15])[0]
            debtor_code = random.choice(agent_debtor[agent])
            debtor_name = DEBTORS[debtor_code]
            day = random.randint(1, 28)
            date = datetime(year, month, day)
            item_code = random.choice(list(ITEMS.keys()))
            desc, unit, base_price = ITEMS[item_code]
            qty = random.randint(1, 20)
            unit_price = round(base_price * random.uniform(0.92, 1.12), 2)
            total = round(qty * unit_price, 2)
            # occasional return/credit note
            if random.random() < 0.04:
                qty = -random.randint(1,3); total = round(qty*unit_price, 2)
            rows.append({
                "Doc Date": date.strftime("%d/%m/%Y"),
                "Debtor Code": debtor_code,
                "Debtor Name": debtor_name,
                "Agent": agent,
                "Item Code": item_code,
                "Detail Description": desc,
                "Qty": qty,
                "Unit Price": unit_price,
                "Total": total,
            })

df = pd.DataFrame(rows)
df.to_excel("invoice_data.xlsx", index=False)
print(f"✅ Generated {len(df)} rows → invoice_data.xlsx")
print(f"   Years: 2020–2025 | Agents: {len(AGENTS)} | Customers: {len(DEBTORS)} | Items: {len(ITEMS)}")
