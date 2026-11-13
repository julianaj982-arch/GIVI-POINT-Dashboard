import os
import datetime
from flask import Flask, render_template, jsonify, request
import pandas as pd

app = Flask(__name__)

# NOTE: File ini adalah versi PORTFOLIO.
# Data asli diganti dengan data dummy untuk keamanan.
EXCEL_FILE = 'data_dummy.xlsx'

def read_sheet(sheet_name, header_row=0):
    if not os.path.exists(EXCEL_FILE):
        return None
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=header_row)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        print(f"Info: Sheet '{sheet_name}' gagal dibaca: {e}")
        return None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/customer-root-cause')
def customer_root_cause_page():
    return render_template('customer_root_cause.html')

@app.route('/forecast')
def forecast_page():
    return render_template('forecast.html')

@app.route('/api/dashboard-data', methods=['GET'])
def get_dashboard_data():
    selected_month = (request.args.get('month') or request.args.get('period') or 'October').strip().capitalize()
    months_list = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

    df_sales = read_sheet('Sales_Data')
    mom_2026, mom_2025, mom_diff, target_val, ach_rate = 0, 0, 0, 0, 0
    ytd_2026, ytd_2025, ytd_diff, ytd_growth = 0, 0, 0, 0
    avg_2026, avg_2025 = 0, 0
    s_2025, s_2026, s_grp = [], [], []

    if df_sales is not None and not df_sales.empty:
        try:
            m_col = [c for c in df_sales.columns if 'month' in c.lower() or 'bulan' in c.lower()]
            month_key = m_col[0] if m_col else df_sales.columns[0]
            row = df_sales[df_sales[month_key].astype(str).str.strip().str.capitalize() == selected_month]
            if not row.empty:
                mom_2026 = float(row['MOM_2026'].values[0]) if 'MOM_2026' in row.columns and pd.notnull(row['MOM_2026'].values[0]) else 0
                mom_2025 = float(row['MOM_2025'].values[0]) if 'MOM_2025' in row.columns and pd.notnull(row['MOM_2025'].values[0]) else 0
                mom_diff = float(row['Incremental Rev'].values[0]) if 'Incremental Rev' in row.columns and pd.notnull(row['Incremental Rev'].values[0]) else (mom_2026 - mom_2025)
                target_val = float(row['Target'].values[0]) if 'Target' in row.columns and pd.notnull(row['Target'].values[0]) else 0
                ach_raw = row['Achivement'].values[0] if 'Achivement' in row.columns else (row['Achievement'].values[0] if 'Achievement' in row.columns else 0)
                ach_rate = float(ach_raw * 100) if pd.notnull(ach_raw) and ach_raw <= 1 else (float(ach_raw) if pd.notnull(ach_raw) else 0)
                ytd_2026 = float(row['YTD_2026'].values[0]) if 'YTD_2026' in row.columns and pd.notnull(row['YTD_2026'].values[0]) else 0
                ytd_2025 = float(row['YTD_2025'].values[0]) if 'YTD_2025' in row.columns and pd.notnull(row['YTD_2025'].values[0]) else 0
                ytd_diff = float(row['TD_Incremental'].values[0]) if 'TD_Incremental' in row.columns and pd.notnull(row['TD_Incremental'].values[0]) else (ytd_2026 - ytd_2025)
                gr2 = row['Growth2'].values[0] if 'Growth2' in row.columns else (row['Growth'].values[0] if 'Growth' in row.columns else 0)
                ytd_growth = float(gr2 * 100) if pd.notnull(gr2) and abs(gr2) <= 1 else (float(gr2) if pd.notnull(gr2) else 0)
                avg_2025 = float(row['Average 2025'].values[0]) if 'Average 2025' in row.columns and pd.notnull(row['Average 2025'].values[0]) else 0
                avg_2026 = float(row['Average 2026'].values[0]) if 'Average 2026' in row.columns and pd.notnull(row['Average 2026'].values[0]) else 0

            for m in months_list:
                m_row = df_sales[df_sales[month_key].astype(str).str.strip().str.capitalize() == m]
                if not m_row.empty:
                    v25 = m_row['MOM_2025'].values[0] if 'MOM_2025' in m_row.columns else None
                    v26 = m_row['MOM_2026'].values[0] if 'MOM_2026' in m_row.columns else None
                    g_val = m_row['Growth'].values[0] if 'Growth' in m_row.columns else None
                    val_2025_m = float(v25) if pd.notnull(v25) and v25 != 0 else None
                    val_2026_m = float(v26) if pd.notnull(v26) and v26 != 0 else None
                    s_2025.append(round(val_2025_m / 1000000, 2) if val_2025_m else None)
                    s_2026.append(round(val_2026_m / 1000000, 2) if val_2026_m else None)
                    if pd.notnull(g_val) and val_2026_m is not None:
                        s_grp.append(round(g_val * 100, 1) if abs(g_val) <= 1 else round(g_val, 1))
                    else:
                        s_grp.append(None)
                else:
                    s_2025.append(None)
                    s_2026.append(None)
                    s_grp.append(None)
        except Exception as e:
            print("Error Sales_Data:", e)

    valid_growth = [g for g in s_grp if g is not None]
    avg_annual_growth = round(sum(valid_growth) / len(valid_growth), 1) if valid_growth else 0.0

    df_traffic = read_sheet('Traffic_KPI')
    qty_val, traffic_val, orders_val, kpi_rate, atv_val, upt_val, csat_val = 0, 0, 0, "0%", "IDR 0", 0.0, "4.8/5"
    if df_traffic is not None and not df_traffic.empty:
        try:
            m_col = [c for c in df_traffic.columns if 'month' in c.lower() or 'bulan' in c.lower()]
            m_key = m_col[0] if m_col else df_traffic.columns[0]
            t_row = df_traffic[df_traffic[m_key].astype(str).str.strip().str.capitalize() == selected_month]
            if not t_row.empty:
                qty_val = int(t_row['Qty'].values[0]) if 'Qty' in t_row.columns and pd.notnull(t_row['Qty'].values[0]) else 0
                traffic_val = int(t_row['Foot Walking'].values[0]) if 'Foot Walking' in t_row.columns and pd.notnull(t_row['Foot Walking'].values[0]) else 0
                conv_col = 'Conversion' if 'Conversion' in t_row.columns else ('Convertion' if 'Convertion' in t_row.columns else '')
                orders_val = int(t_row[conv_col].values[0]) if conv_col and pd.notnull(t_row[conv_col].values[0]) else 0
                kpi_raw = t_row['KPI'].values[0] if 'KPI' in t_row.columns else 0
                kpi_rate = f"{float(kpi_raw)*100:.0f}%" if pd.notnull(kpi_raw) and float(kpi_raw) <= 1 else f"{float(kpi_raw):.0f}%"
                atv_raw = t_row['ATV'].values[0] if 'ATV' in t_row.columns else 0
                atv_val = f"IDR {float(atv_raw):,.0f}".replace(",", ".") if pd.notnull(atv_raw) else "IDR 0"
                upt_val = float(t_row['UPT'].values[0]) if 'UPT' in t_row.columns and pd.notnull(t_row['UPT'].values[0]) else 0.0
                csat_val = str(t_row['CSAT'].values[0] if 'CSAT' in t_row.columns and pd.notnull(t_row['CSAT'].values[0]) else '4.8/5')
        except Exception as e:
            print("Error Traffic_KPI:", e)

    df_weekly = read_sheet('Weekly_Sales', header_row=2)
    weekly_sales = []
    weekly_total_amt = 0
    weekly_growth_str = "0.0%"
    if df_weekly is not None and not df_weekly.empty:
        try:
            m_col = [c for c in df_weekly.columns if 'month' in c.lower() or 'bulan' in c.lower()]
            m_key = m_col[0] if m_col else df_weekly.columns[0]
            w_row = df_weekly[df_weekly[m_key].astype(str).str.strip().str.capitalize() == selected_month]
            if not w_row.empty:
                w1 = float(w_row['W1'].values[0]) if 'W1' in w_row.columns and pd.notnull(w_row['W1'].values[0]) else 0
                w2 = float(w_row['W2'].values[0]) if 'W2' in w_row.columns and pd.notnull(w_row['W2'].values[0]) else 0
                w3 = float(w_row['W3'].values[0]) if 'W3' in w_row.columns and pd.notnull(w_row['W3'].values[0]) else 0
                w4 = float(w_row['W4'].values[0]) if 'W4' in w_row.columns and pd.notnull(w_row['W4'].values[0]) else 0
                w5 = float(w_row['W5'].values[0]) if 'W5' in w_row.columns and pd.notnull(w_row['W5'].values[0]) else 0
                weekly_total_amt = float(w_row['TOTAL'].values[0]) if 'TOTAL' in w_row.columns and pd.notnull(w_row['TOTAL'].values[0]) else (w1 + w2 + w3 + w4 + w5)
                tot_calc = weekly_total_amt if weekly_total_amt > 0 else 1
                if 'Growth' in w_row.columns and pd.notnull(w_row['Growth'].values[0]):
                    gr_w = float(w_row['Growth'].values[0])
                    weekly_growth_str = f"{gr_w*100:.1f}%" if abs(gr_w) <= 1 else f"{gr_w:.1f}%"
                weekly_sales = [
                    {"week": "W1", "amount": w1, "amount_str": f"IDR {w1:,.0f}".replace(",", "."), "percentage": (w1/tot_calc)*100, "pct_str": f"{(w1/tot_calc)*100:.1f}%"},
                    {"week": "W2", "amount": w2, "amount_str": f"IDR {w2:,.0f}".replace(",", "."), "percentage": (w2/tot_calc)*100, "pct_str": f"{(w2/tot_calc)*100:.1f}%"},
                    {"week": "W3", "amount": w3, "amount_str": f"IDR {w3:,.0f}".replace(",", "."), "percentage": (w3/tot_calc)*100, "pct_str": f"{(w3/tot_calc)*100:.1f}%"},
                    {"week": "W4", "amount": w4, "amount_str": f"IDR {w4:,.0f}".replace(",", "."), "percentage": (w4/tot_calc)*100, "pct_str": f"{(w4/tot_calc)*100:.1f}%"},
                    {"week": "W5", "amount": w5, "amount_str": f"IDR {w5:,.0f}".replace(",", "."), "percentage": (w5/tot_calc)*100, "pct_str": f"{(w5/tot_calc)*100:.1f}%"}
                ]
        except Exception as e:
            print("Error Weekly_Sales:", e)

    df_cat = read_sheet('Category')
    categories = []
    top_category_name = "-"
    top_category_pct = "0%"
    if df_cat is not None and not df_cat.empty:
        try:
            m_col = [c for c in df_cat.columns if 'month' in c.lower() or 'bulan' in c.lower()]
            m_key = m_col[0] if m_col else df_cat.columns[0]
            c_row = df_cat[df_cat[m_key].astype(str).str.strip().str.capitalize() == selected_month]
            if not c_row.empty:
                cat_cols = [c for c in df_cat.columns if c not in [m_key, 'Total']]
                tot_val = float(c_row['Total'].values[0]) if 'Total' in c_row.columns and pd.notnull(c_row['Total'].values[0]) else sum([float(c_row[c].values[0]) for c in cat_cols if pd.notnull(c_row[c].values[0])])
                if tot_val == 0: tot_val = 1
                raw_categories = []
                for col in cat_cols:
                    val = float(c_row[col].values[0]) if pd.notnull(c_row[col].values[0]) else 0
                    pct_raw = (val / tot_val) * 100
                    raw_categories.append({"name": col, "amount": val, "amount_str": f"IDR {val:,.0f}".replace(",", "."), "pct_num": round(pct_raw, 1), "percentage": f"{pct_raw:.1f}%"})
                categories = sorted(raw_categories, key=lambda x: x['amount'], reverse=True)
                if categories and categories[0]['amount'] > 0:
                    top_category_name = categories[0]['name']
                    top_category_pct = categories[0]['percentage']
        except Exception as e:
            print("Error Category:", e)

    df_rev = read_sheet('Revenue_Source')
    channels = {"offline": {"amount": "IDR 0", "percentage": "0.0%"}, "website": {"amount": "IDR 0", "percentage": "0.0%"}, "event": {"amount": "IDR 0", "percentage": "0.0%"}}
    if df_rev is not None and not df_rev.empty:
        try:
            m_col = [c for c in df_rev.columns if 'period' in c.lower() or 'month' in c.lower() or 'bulan' in c.lower()]
            m_key = m_col[0] if m_col else df_rev.columns[0]
            r_row = df_rev[df_rev[m_key].astype(str).str.strip().str.capitalize() == selected_month]
            if not r_row.empty:
                cols = [c for c in df_rev.columns if c != m_key and c.lower() != 'total']
                off_val = float(r_row[cols[0]].values[0]) if len(cols) > 0 and pd.notnull(r_row[cols[0]].values[0]) else 0
                web_val = float(r_row[cols[1]].values[0]) if len(cols) > 1 and pd.notnull(r_row[cols[1]].values[0]) else 0
                evt_val = float(r_row[cols[2]].values[0]) if len(cols) > 2 and pd.notnull(r_row[cols[2]].values[0]) else 0
                tot_rev = off_val + web_val + evt_val if (off_val + web_val + evt_val) > 0 else 1
                channels = {
                    "offline": {"amount": f"IDR {off_val:,.0f}".replace(",", "."), "percentage": f"{(off_val/tot_rev)*100:.1f}%"},
                    "website": {"amount": f"IDR {web_val:,.0f}".replace(",", "."), "percentage": f"{(web_val/tot_rev)*100:.1f}%"},
                    "event": {"amount": f"IDR {evt_val:,.0f}".replace(",", "."), "percentage": f"{(evt_val/tot_rev)*100:.1f}%"}
                }
        except Exception as e:
            print("Error Revenue_Source:", e)

    df_trend = read_sheet('Sales_Trend')
    top_5_grid = {"top_case": {"unit": "0 Unit", "items": [], "total": "IDR 0"}, "racking": {"unit": "0 Unit", "items": [], "total": "IDR 0"}, "soft_bag": {"unit": "0 Unit", "items": [], "total": "IDR 0"}, "accessories": {"unit": "0 Unit", "items": [], "total": "IDR 0"}, "fitting_kits": {"unit": "0 Unit", "items": [], "total": "IDR 0"}}
    if df_trend is not None and not df_trend.empty:
        try:
            for cat_key, target_cat in [("top_case", "Top Case"), ("racking", "Racking"), ("soft_bag", "Soft Bag"), ("accessories", "Accessories"), ("fitting_kits", "Fitting Kits")]:
                sub = df_trend[df_trend['Category'].astype(str).str.strip().str.lower() == target_cat.lower()]
                if not sub.empty:
                    tot_qty = sub['QTY'].sum() if 'QTY' in sub.columns else 0
                    tot_rev = sub['Revenue'].sum() if 'Revenue' in sub.columns else 0
                    items_list = []
                    for i, (_, r) in enumerate(sub.head(5).iterrows()):
                        sku_name = str(r.get('SKU', ''))
                        q_val = float(r.get('QTY', 0)) if pd.notnull(r.get('QTY', 0)) else 0
                        rev_val = float(r.get('Revenue', 0)) if pd.notnull(r.get('Revenue', 0)) else 0
                        items_list.append({"rank": i + 1, "sku": sku_name, "qty": f"{int(q_val)} Unit", "revenue": f"IDR {rev_val:,.0f}".replace(",", ".")})
                    top_5_grid[cat_key] = {"unit": f"{int(tot_qty)} Unit", "items": items_list, "total": f"IDR {tot_rev:,.0f}".replace(",", ".")}
        except Exception as e:
            print("Error Sales_Trend:", e)

    mom_label = "Incremental Revenue" if mom_diff >= 0 else "Discremental"
    ytd_label = "Incremental Revenue" if ytd_diff >= 0 else "Discremental"
    mom_pct_val = ((mom_2026 - mom_2025) / mom_2025) * 100 if mom_2025 > 0 else 0
    avg_diff = avg_2026 - avg_2025
    avg_label = "Inc Average" if avg_diff >= 0 else "Dec Average"
    avg_pct_val = ((avg_2026 - avg_2025) / avg_2025) * 100 if avg_2025 > 0 else 0

    return jsonify({
        "status": "success", "selected_month": selected_month, "available_months": months_list,
        "last_updated": datetime.datetime.now().strftime("%d %b %Y %H:%M"),
        "kpi": {"traffic": f"{traffic_val:,}".replace(",", "."), "transactions": f"{orders_val:,}".replace(",", "."), "conversion_rate": kpi_rate, "qty_sales": f"{qty_val:,}".replace(",", "."), "atv": atv_val, "upt": f"{upt_val:.1f}", "csat": csat_val},
        "mom": {"rev_2026": f"IDR {mom_2026:,.0f}".replace(",", "."), "rev_2025": f"IDR {mom_2025:,.0f}".replace(",", "."), "diff_amount": f"IDR {mom_diff:,.0f}".replace(",", "."), "diff_label": mom_label, "target": f"IDR {target_val:,.0f}".replace(",", "."), "ach_rate": f"{ach_rate:.1f}%", "diff_percent": f"{mom_pct_val:.1f}%", "is_negative": mom_pct_val < 0},
        "ytd": {"rev_2026": f"IDR {ytd_2026:,.0f}".replace(",", "."), "rev_2025": f"IDR {ytd_2025:,.0f}".replace(",", "."), "diff_amount": f"IDR {ytd_diff:,.0f}".replace(",", "."), "diff_label": ytd_label, "growth": f"+{ytd_growth:.1f}%" if ytd_growth >= 0 else f"{ytd_growth:.1f}%", "is_negative": ytd_growth < 0},
        "sales_average": {"rev_2026": f"IDR {avg_2026:,.0f}".replace(",", "."), "rev_2025": f"IDR {avg_2025:,.0f}".replace(",", "."), "diff_amount": f"IDR {avg_diff:,.0f}".replace(",", "."), "diff_label": avg_label, "diff_percent": f"{avg_pct_val:.1f}%", "is_negative": avg_pct_val < 0},
        "annual_chart": {"months": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], "full_months": months_list, "sales_2025": s_2025, "sales_2026": s_2026, "growth": s_grp, "avg_annual_growth": f"{avg_annual_growth:.1f}%"},
        "weekly_sales": {"total_str": f"IDR {weekly_total_amt:,.0f}".replace(",", "."), "growth_str": weekly_growth_str, "weeks": weekly_sales},
        "category_meta": {"top_category": top_category_name, "top_percentage": top_category_pct, "total_categories": len(categories)},
        "categories": categories, "channels": channels, "top_5_grid": top_5_grid
    })

@app.route('/api/customer-root-cause', methods=['GET'])
def get_customer_root_cause():
    # VERSI ANONIM: Nama customer asli dihapus, hanya tampil agregat untuk portofolio
    df_rc = read_sheet('Customer_Root_Cause')
    records = []
    categories_summary = {}
    brands_summary = {}
    details_summary = {}
    if df_rc is not None and not df_rc.empty:
        try:
            df_rc = df_rc.dropna(how='all')
            for idx, row in df_rc.iterrows():
                # Anonimisasi: tidak menampilkan nama customer
                customer_id = f"CUST-{idx+1:03d}"
                raw_date = row.get('Date', '')
                date_str = str(raw_date)[:10] if pd.notnull(raw_date) else '-'
                if isinstance(raw_date, (datetime.datetime, pd.Timestamp)):
                    date_str = raw_date.strftime('%d-%b-%Y')
                brand = str(row.get('Brand Motorcycle', '-')).strip().title()
                m_type = str(row.get('Type', '-')).strip()
                r_code = str(row.get('Reason Code', '-')).strip()
                cat = str(row.get('Root Cause Category', '-')).strip().title()
                detail = str(row.get('Root Cause Detail', '-')).strip()
                action = str(row.get('Follow-up Action', '-')).strip()
                notes = str(row.get('Notes', '-')).strip()
                # Notes disamarkan jika mengandung info pribadi
                if len(notes) > 20:
                    notes = "Follow-up by team"
                records.append({"date": date_str, "customer": customer_id, "brand": brand, "type": m_type, "reason_code": r_code, "category": cat, "detail": detail, "action": action, "notes": notes})
                if cat != '-': categories_summary[cat] = categories_summary.get(cat, 0) + 1
                if brand != '-': brands_summary[brand] = brands_summary.get(brand, 0) + 1
                if detail != '-': details_summary[detail] = details_summary.get(detail, 0) + 1
        except Exception as e:
            print("Error Customer_Root_Cause:", e)
    top_cat = max(categories_summary, key=categories_summary.get) if categories_summary else "-"
    top_detail = max(details_summary, key=details_summary.get) if details_summary else "-"
    top_brand = max(brands_summary, key=brands_summary.get) if brands_summary else "-"
    return jsonify({"status": "success", "last_updated": datetime.datetime.now().strftime("%d %b %Y %H:%M"), "total_records": len(records), "kpi": {"top_category": top_cat, "top_detail": top_detail, "top_brand": top_brand}, "categories_breakdown": [{"name": k, "count": v} for k, v in categories_summary.items()], "brands_breakdown": [{"name": k, "count": v} for k, v in brands_summary.items()], "records": records})

@app.route('/api/forecast-data', methods=['GET'])
def get_forecast_data():
    try:
        # 1. COBA BACA DARI Traffic_KPI DULU, KALAU GAK ADA BARU Sales_Data
        df_sales = read_sheet('Traffic_KPI')
        if df_sales is None or df_sales.empty:
            df_sales = read_sheet('Sales_Data')
        if df_sales is None or df_sales.empty:
            return jsonify({"status": "error", "message": "Sheet Traffic_KPI / Sales_Data tidak ketemu"})

        # 2. CARI KOLOM BULAN & SALES OTOMATIS (gak peduli namanya Month/Bulan, Sales/MOM_2026)
        cols_lower = {c.lower(): c for c in df_sales.columns}
        m_key = next((cols_lower[k] for k in cols_lower if 'month' in k or 'bulan' in k), df_sales.columns[0])
        s_key = next((cols_lower[k] for k in cols_lower if 'sales' in k or 'mom_2026' in k or 'actual' in k), df_sales.columns[1])

        hist_months = ["January","February","March","April","May","June","July","August","September","October"]
        historical_sales = []
        for m in hist_months:
            m_row = df_sales[df_sales[m_key].astype(str).str.strip().str.capitalize() == m]
            if not m_row.empty:
                v = m_row[s_key].values[0]
                # Bersihin titik / koma biar "209.885.207" jadi 209885207
                if isinstance(v, str):
                    v = v.replace('.','').replace(',','').replace('IDR','').strip()
                historical_sales.append(float(v) if pd.notnull(v) and str(v) != '' else 0)
            else:
                historical_sales.append(0)

        # 3. HITUNG FORECAST Q4 DARI DATA ASLI (bukan hardcode)
        # Okt = data asli Okt, Nov = rata-rata 3 bulan terakhir + 5%, Des = Nov + 5%
        oct_actual = historical_sales[9] if len(historical_sales) > 9 else 0
        avg_last3 = sum(historical_sales[7:10]) / 3 if sum(historical_sales[7:10]) > 0 else oct_actual

        oct_base, nov_base, dec_base = oct_actual, avg_last3 * 1.05, avg_last3 * 1.05 * 1.05
        oct_opt, nov_opt, dec_opt = oct_actual, avg_last3 * 1.12, avg_last3 * 1.12 * 1.12
        oct_pess, nov_pess, dec_pess = oct_actual, avg_last3 * 0.95, avg_last3 * 0.95 * 0.95

        actual_ytd = sum(historical_sales)
        # ... (bagian bawah yang lu screenshot tadi, biarin aja, tinggal return)

        return jsonify({
            "status": "success",
            "last_updated": datetime.datetime.now().strftime("%d %b %Y %H:%M"),
            "historical_sales": historical_sales,
            "hist_months": hist_months,
            "actual_ytd_total": actual_ytd,
            "scenarios": {
                "base": {"oct": oct_base, "nov": nov_base, "dec": dec_base},
                "optimistic": {"oct": oct_opt, "nov": nov_opt, "dec": dec_opt},
                "pessimistic": {"oct": oct_pess, "nov": nov_pess, "dec": dec_pess}
            }
        })
    except Exception as e:
        print(f"FORECAST ERROR: {e}")
        return jsonify({"status": "error", "message": str(e)})
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5003))
    app.run(host='0.0.0.0', port=port)