import pandas as pd
import random
from datetime import datetime

random.seed(42)
months = ["January","February","March","April","May","June","July","August","September","October","November","December"]

# 1. Sales_Data
sales_rows = []
ytd26 = ytd25 = 0
for i, m in enumerate(months):
    mom26 = random.randint(180_000_000, 320_000_000)
    mom25 = random.randint(150_000_000, 280_000_000)
    ytd26 += mom26
    ytd25 += mom25
    target = int(mom26 * 1.08)
    ach = mom26 / target if target else 0
    growth = (mom26 - mom25) / mom25 if mom25 else 0
    sales_rows.append({
        "Month": m, "MOM_2026": mom26, "MOM_2025": mom25,
        "Incremental Rev": mom26 - mom25, "Target": target, "Achivement": ach,
        "YTD_2026": ytd26, "YTD_2025": ytd25, "TD_Incremental": ytd26 - ytd25,
        "Growth2": growth, "Growth": growth,
        "Average 2025": ytd25 // (i+1), "Average 2026": ytd26 // (i+1)
    })
df_sales = pd.DataFrame(sales_rows)

# 2. Traffic_KPI
traffic_rows = []
for m in months:
    qty = random.randint(800, 1500)
    foot = random.randint(4000, 9000)
    conv = int(foot * random.uniform(0.12, 0.22))
    traffic_rows.append({
        "Month": m, "Qty": qty, "Foot Walking": foot, "Conversion": conv,
        "KPI": random.uniform(0.75, 0.98), "ATV": random.randint(850000, 1450000),
        "UPT": round(random.uniform(1.2, 2.4),1), "CSAT": "4.8/5"
    })
df_traffic = pd.DataFrame(traffic_rows)

# 3. Weekly_Sales - header di baris ke-3 (header_row=2)
weekly_rows = []
for m in months:
    w = [random.randint(35_000_000, 70_000_000) for _ in range(5)]
    total = sum(w)
    weekly_rows.append({"Month": m, "W1": w[0], "W2": w[1], "W3": w[2], "W4": w[3], "W5": w[4], "TOTAL": total, "Growth": random.uniform(-0.05, 0.15)})
df_weekly = pd.DataFrame(weekly_rows)

# 4. Category
cat_rows = []
for m in months:
    top = random.randint(60_000_000, 110_000_000)
    rack = random.randint(30_000_000, 60_000_000)
    soft = random.randint(25_000_000, 50_000_000)
    acc = random.randint(15_000_000, 35_000_000)
    fit = random.randint(10_000_000, 25_000_000)
    cat_rows.append({"Month": m, "Top Case": top, "Racking": rack, "Soft Bag": soft, "Accessories": acc, "Fitting Kits": fit, "Total": top+rack+soft+acc+fit})
df_cat = pd.DataFrame(cat_rows)

# 5. Revenue_Source
rev_rows = []
for m in months:
    off = random.randint(120_000_000, 220_000_000)
    web = random.randint(20_000_000, 50_000_000)
    evt = random.randint(5_000_000, 25_000_000)
    rev_rows.append({"Period": m, "Offline": off, "Website": web, "Event": evt})
df_rev = pd.DataFrame(rev_rows)

# 6. Sales_Trend
skus = {
    "Top Case": ["B47NML Atlas", "B33NML Atlas", "E43NML", "B42N Antartica", "TRK52N"],
    "Racking": ["SR2159", "SR2146", "PLX2146", "SRA2146", "SR1171"],
    "Soft Bag": ["EA101B", "EA106B", "XS308", "ST609", "EA122"],
    "Accessories": ["S150 Visor", "Backrest E43", "Top Case Pad", "Lock Set", "Inner Bag"],
    "Fitting Kits": ["KIT Yamaha Nmax", "KIT Honda PCX", "KIT Vespa", "KIT ADV160", "KIT Xmax"]
}
trend_rows = []
for cat, items in skus.items():
    for sku in items:
        trend_rows.append({"Category": cat, "SKU": sku, "QTY": random.randint(20, 120), "Revenue": random.randint(8_000_000, 45_000_000)})
df_trend = pd.DataFrame(trend_rows)

# 7. Customer_Root_Cause
brands = ["Honda", "Yamaha", "Vespa", "Kawasaki", "Suzuki"]
cats_rc = ["Price", "Product Fitment", "Availability", "Service", "Product Knowledge"]
details_rc = ["Price too high vs competitor", "Fitting kit not available", "Stock empty", "Need installation help", "Confused about size"]
rc_rows = []
for i in range(60):
    rc_rows.append({
        "Date": datetime(2026, random.randint(1,10), random.randint(1,28)),
        "Brand Motorcycle": random.choice(brands), "Type": random.choice(["Nmax","PCX","Vario","ADV160","Xmax"]),
        "Reason Code": f"RC-{random.randint(1,5):02d}", "Root Cause Category": random.choice(cats_rc),
        "Root Cause Detail": random.choice(details_rc), "Follow-up Action": "Follow-up by team", "Notes": "Follow-up by team"
    })
df_rc = pd.DataFrame(rc_rows)

# Simpan
with pd.ExcelWriter("data_dummy.xlsx", engine="openpyxl") as writer:
    df_sales.to_excel(writer, sheet_name="Sales_Data", index=False)
    df_traffic.to_excel(writer, sheet_name="Traffic_KPI", index=False)
    # Weekly sengaja dikasih 2 baris judul di atas biar header_row=2 cocok
    df_weekly.to_excel(writer, sheet_name="Weekly_Sales", index=False, startrow=2)
    df_cat.to_excel(writer, sheet_name="Category", index=False)
    df_rev.to_excel(writer, sheet_name="Revenue_Source", index=False)
    df_trend.to_excel(writer, sheet_name="Sales_Trend", index=False)
    df_rc.to_excel(writer, sheet_name="Customer_Root_Cause", index=False)

print("BERHASIL! data_dummy.xlsx udah kebikin!")