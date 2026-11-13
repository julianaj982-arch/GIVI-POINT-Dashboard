# fix_dummy.py - V4 FIX Traffic dtype
import pandas as pd
import numpy as np
import random

INPUT_FILE = "data_dummy.xlsx"
OUTPUT_FILE = "data_dummy_FIXED.xlsx"
random.seed(42)
np.random.seed(42)

def load_sheet(xls, sheet_name):
    raw = pd.read_excel(xls, sheet_name, header=None)
    header_row = 0
    for i in range(min(5, len(raw))):
        row_vals = [str(x).lower() for x in raw.iloc[i].tolist()]
        if any("month" in v or "period" in v for v in row_vals):
            header_row = i
            break
    print(f"[{sheet_name}] header di baris {header_row+1}")
    df = pd.read_excel(xls, sheet_name, header=header_row)
    df = df.loc[:, ~df.columns.astype(str).str.contains('Unnamed')]
    return df

def get_month_col(df):
    for c in df.columns:
        if str(c).lower().strip() in ["month", "period", "bulan"]:
            return c
    return None

xls = pd.ExcelFile(INPUT_FILE)
print(f"Sheet: {xls.sheet_names}")

# 1. SALES_DATA
df_sales = load_sheet(xls, "Sales_Data")
mcol_sales = get_month_col(df_sales)
patokan = dict(zip(df_sales[mcol_sales], df_sales["MOM_2026"]))
df_sales["Achievement"] = df_sales["MOM_2026"] / df_sales["Target"]
df_sales["Growth_YoY"] = (df_sales["MOM_2026"] - df_sales["MOM_2025"]) / df_sales["MOM_2025"]
df_sales["YTD_2026"] = df_sales["MOM_2026"].cumsum()
df_sales["YTD_2025"] = df_sales["MOM_2025"].cumsum()
df_sales["YTD_Target"] = df_sales["Target"].cumsum()
print("[OK] Sales_Data")

# 2. WEEKLY
if "Weekly_Sales" in xls.sheet_names:
    df_weekly = load_sheet(xls, "Weekly_Sales")
    mcol = get_month_col(df_weekly)
    week_cols = [c for c in df_weekly.columns if str(c).strip() in ["W1","W2","W3","W4","W5"]]
    df_weekly = df_weekly.dropna(subset=[mcol])
    for idx, row in df_weekly.iterrows():
        target = patokan.get(row[mcol])
        if target is None: continue
        weights = np.array([0.18,0.20,0.22,0.22,0.18][:len(week_cols)])
        weights = weights + np.random.uniform(-0.02,0.02,size=len(weights))
        weights = weights/weights.sum()
        new_vals = (weights*target).round(0)
        new_vals[2] += target - new_vals.sum()
        for i, col in enumerate(week_cols):
            df_weekly.at[idx, col] = new_vals[i]
    if "TOTAL" in df_weekly.columns:
        df_weekly["TOTAL"] = df_weekly[week_cols].sum(axis=1)
    if "Growth" in df_weekly.columns:
        df_weekly["Growth"] = df_weekly["TOTAL"].pct_change()
    print("[OK] Weekly_Sales")

# 3. CATEGORY
if "Category" in xls.sheet_names:
    df_cat = load_sheet(xls, "Category")
    mcol = get_month_col(df_cat)
    cat_cols = [c for c in df_cat.columns if c not in [mcol,"TOTAL","Total","Growth"]]
    df_cat = df_cat.dropna(subset=[mcol])
    for idx, row in df_cat.iterrows():
        target = patokan.get(row[mcol])
        if target is None: continue
        n=len(cat_cols)
        base = np.array([0.38,0.22,0.15,0.13,0.12][:5]) if n>=5 else np.array([1/n]*n)
        if n>5:
            base = np.concatenate([base*0.85, np.array([0.05]*(n-5))])
            base = base/base.sum()
        base = base + np.random.uniform(-0.02,0.02,size=n)
        base = np.clip(base,0.02,None); base=base/base.sum()
        new_vals=(base*target).round(0)
        new_vals[0]+=target-new_vals.sum()
        for i,col in enumerate(cat_cols):
            df_cat.at[idx,col]=new_vals[i]
    if "Total" in df_cat.columns:
        df_cat["Total"]=df_cat[cat_cols].sum(axis=1)
    print("[OK] Category")

# 4. REVENUE
if "Revenue_Source" in xls.sheet_names:
    df_rev = load_sheet(xls, "Revenue_Source")
    mcol = get_month_col(df_rev)
    rev_cols = [c for c in df_rev.columns if c not in [mcol,"TOTAL","Total","Growth"]]
    df_rev = df_rev.dropna(subset=[mcol])
    for idx, row in df_rev.iterrows():
        target = patokan.get(row[mcol])
        if target is None: continue
        n=len(rev_cols)
        weights = np.array([0.62,0.28,0.10]) if n==3 else np.array([1/n]*n)
        weights = weights + np.random.uniform(-0.02,0.02,size=n)
        weights=np.clip(weights,0.02,None); weights=weights/weights.sum()
        new_vals=(weights*target).round(0)
        new_vals[0]+=target-new_vals.sum()
        for i,col in enumerate(rev_cols):
            df_rev.at[idx,col]=new_vals[i]
    if "TOTAL" in df_rev.columns:
        df_rev["TOTAL"]=df_rev[rev_cols].sum(axis=1)
    if "Total" in df_rev.columns:
        df_rev["Total"]=df_rev[rev_cols].sum(axis=1)
    print("[OK] Revenue_Source")

# 5. SALES_TREND
if "Sales_Trend" in xls.sheet_names:
    df_trend = load_sheet(xls, "Sales_Trend")
    mcol = get_month_col(df_trend)
    if mcol and "MOM_2026" in df_trend.columns:
        for idx,row in df_trend.iterrows():
            if row[mcol] in patokan:
                df_trend.at[idx,"MOM_2026"]=patokan[row[mcol]]
    print("[OK] Sales_Trend")

# 6. TRAFFIC_KPI - V4 FIX DTYPE
if "Traffic_KPI" in xls.sheet_names:
    df_traffic = load_sheet(xls, "Traffic_KPI")
    mcol = get_month_col(df_traffic)
    print(f"Traffic cols: {list(df_traffic.columns)}")

    # === INI FIX-NYA, 3 BARIS INI YANG BIKIN GAK ERROR LAGI ===
    # Paksa kolom jadi tipe yang bisa nampung desimal / teks
    if "Conversion" in df_traffic.columns:
        df_traffic["Conversion"] = df_traffic["Conversion"].astype(float)
    if "KPI" in df_traffic.columns:
        df_traffic["KPI"] = df_traffic["KPI"].astype(float)
    if "ATV" in df_traffic.columns:
        df_traffic["ATV"] = df_traffic["ATV"].astype(float)
    if "CSAT" in df_traffic.columns:
        df_traffic["CSAT"] = df_traffic["CSAT"].astype(str)
    if "Qty" in df_traffic.columns:
        df_traffic["Qty"] = df_traffic["Qty"].astype(float)
    if "Foot Walking" in df_traffic.columns:
        df_traffic["Foot Walking"] = df_traffic["Foot Walking"].astype(float)
    # =========================================================

    if mcol:
        df_traffic=df_traffic.dropna(subset=[mcol])
        for idx,row in df_traffic.iterrows():
            target=patokan.get(row[mcol])
            if target is None: continue
            atv=random.uniform(1800000,2600000)
            qty=round(target/atv)
            conv=random.uniform(0.18,0.28)
            foot=round(qty/conv)
            csat_val=round(random.uniform(4.6,4.9),1)
            if "ATV" in df_traffic.columns: df_traffic.at[idx,"ATV"]=round(atv,0)
            if "Qty" in df_traffic.columns: df_traffic.at[idx,"Qty"]=qty
            for c in ["Foot Walking","Foot_Walking","Footfall"]:
                if c in df_traffic.columns: df_traffic.at[idx,c]=foot
            if "Conversion" in df_traffic.columns: df_traffic.at[idx,"Conversion"]=round(qty/foot,4)
            if "CSAT" in df_traffic.columns: df_traffic.at[idx,"CSAT"]=f"{csat_val}/5"
            if "Sales" in df_traffic.columns: df_traffic.at[idx,"Sales"]=qty*round(atv,0)
    print("[OK] Traffic_KPI")

# 7. CUSTOMER
if "Customer_Root_Cause" in xls.sheet_names:
    df_cust = load_sheet(xls, "Customer_Root_Cause")
    print(f"Customer cols: {list(df_cust.columns)}")
    fix_map={"Honda - Nmax":"Yamaha - Nmax","Vespa - Nmax":"Yamaha - Nmax","Honda - Aerox":"Yamaha - Aerox","Suzuki - Vario":"Honda - Vario","Yamaha - Vario":"Honda - Vario","Yamaha - PCX":"Honda - PCX","Suzuki - PCX":"Honda - PCX","Honda - Xmax":"Yamaha - Xmax"}
    for col in df_cust.columns:
        if any(k in str(col).lower() for k in ["motor","model","vehicle","bike","unit"]):
            df_cust[col]=df_cust[col].replace(fix_map)
    if "Brand" in df_cust.columns and "Model" in df_cust.columns:
        for idx,row in df_cust.iterrows():
            model=str(row["Model"]).lower()
            if "nmax" in model or "aerox" in model or "xmax" in model: df_cust.at[idx,"Brand"]="Yamaha"
            elif "vario" in model or "pcx" in model or "beat" in model: df_cust.at[idx,"Brand"]="Honda"
    print("[OK] Customer_Root_Cause")

# SAVE
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_sales.to_excel(writer, sheet_name="Sales_Data", index=False)
    if "Weekly_Sales" in xls.sheet_names: df_weekly.to_excel(writer, sheet_name="Weekly_Sales", index=False)
    if "Category" in xls.sheet_names: df_cat.to_excel(writer, sheet_name="Category", index=False)
    if "Revenue_Source" in xls.sheet_names: df_rev.to_excel(writer, sheet_name="Revenue_Source", index=False)
    if "Sales_Trend" in xls.sheet_names: df_trend.to_excel(writer, sheet_name="Sales_Trend", index=False)
    if "Traffic_KPI" in xls.sheet_names: df_traffic.to_excel(writer, sheet_name="Traffic_KPI", index=False)
    if "Customer_Root_Cause" in xls.sheet_names: df_cust.to_excel(writer, sheet_name="Customer_Root_Cause", index=False)

print(f"\n=== SELESAI! File baru: {OUTPUT_FILE} ===")