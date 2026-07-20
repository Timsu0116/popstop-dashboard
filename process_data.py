import pandas as pd
import numpy as np
import os
import glob
import re
import pickle
from datetime import datetime, date

BASE_INPUT = r"D:\Weekly Report Working Folder\Retail_Stock_System\01_Input"
OUTPUT_DIR = r"D:\PopStop_App\data"

SALES_COMPANY_DIR  = os.path.join(BASE_INPUT, "Sales_Company")
SALES_STORE_DIR    = os.path.join(BASE_INPUT, "Sales_Store")
SOH_WEEKLY_DIR      = os.path.join(BASE_INPUT, "SOH_Weekly")
DUSTY_COMPANY_DIR  = os.path.join(BASE_INPUT, "Dusty_Inventory_Company")
DUSTY_STORE_DIR    = os.path.join(BASE_INPUT, "Dusty_Inventory_Store")
PRODUCT_MASTER_DIR = os.path.join(BASE_INPUT, "Product_Master")

os.makedirs(OUTPUT_DIR, exist_ok=True)
TODAY = date.today()

print("=" * 60)
print("PopStop Data Processing")
print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("=" * 60)

def fix_sku(val):
    if pd.isna(val):
        return None
    val = str(val).strip()
    try:
        # Use Decimal for full precision (avoids float rounding on 13-digit barcodes)
        from decimal import Decimal
        return str(int(Decimal(val)))
    except Exception:
        try:
            return str(int(float(val)))
        except Exception:
            return val

def safe_date(val):
    if pd.isna(val):
        return None
    try:
        d = pd.to_datetime(val).date()
        if d < date(2000, 1, 1):
            return None
        return d
    except:
        return None

def days_since(d):
    if d is None:
        return None
    try:
        return (TODAY - d).days
    except:
        return None

def extract_week_start(filename):
    m = re.search(r'\((\d{4}-\d{2}-\d{2})', filename)
    if m:
        return pd.to_datetime(m.group(1)).date()
    return None

# week_key will be assigned after we know the full sorted order of files
# Placeholder: return the date string so we can sort first
def get_week_key_from_index(idx):
    """Returns 2026-W01, 2026-W02 ... based on sorted file order (1-indexed)"""
    return f"2026-W{idx:02d}"

# ============================================================
# 1. Product Master
# ============================================================
print("\n[1/9] Loading Product Master...")
pm_files = glob.glob(os.path.join(PRODUCT_MASTER_DIR, "*.csv"))
pm_file  = sorted(pm_files, key=os.path.getmtime, reverse=True)[0]
pm_raw   = pd.read_csv(pm_file, dtype=str)

# Fix scientific-notation SKUs in Product Master using supplier_code as fallback key.
# When Lightspeed exports PM as CSV, long numeric SKUs get truncated to scientific notation
# (e.g. "6.62E+11"). Since fix_sku(float) loses precision on 13-digit barcodes, we also
# build a supplier_code → supply_price lookup so products can still be joined by supplier_code.

pm = pm_raw[[
    "id","sku","name","product_category","tags",
    "supply_price","retail_price","brand_name","supplier_name",
    "supplier_code","active","track_inventory",
    "inventory_Dunedin","inventory_Office","inventory_Papanui",
    "inventory_Queensgate","inventory_Riccarton","inventory_Richmond",
    "inventory_Sylvia_Park","inventory_Te_Rapa","inventory_WH1",
    "inventory_Warehouse_AKL"
]].copy()

pm.rename(columns={
    "name":"product","product_category":"category","tags":"tag",
    "brand_name":"brand","supplier_name":"supplier",
    "inventory_Dunedin":"soh_Dunedin","inventory_Office":"soh_Office",
    "inventory_Papanui":"soh_Papanui","inventory_Queensgate":"soh_Queensgate",
    "inventory_Riccarton":"soh_Riccarton","inventory_Richmond":"soh_Richmond",
    "inventory_Sylvia_Park":"soh_SylviaPark","inventory_Te_Rapa":"soh_Terapa",
    "inventory_WH1":"soh_WH1","inventory_Warehouse_AKL":"soh_WarehouseAKL"
}, inplace=True)

pm["sku"] = pm["sku"].apply(fix_sku)
soh_cols = ["soh_Dunedin","soh_Office","soh_Papanui","soh_Queensgate",
            "soh_Riccarton","soh_Richmond","soh_SylviaPark","soh_Terapa",
            "soh_WH1","soh_WarehouseAKL"]
for c in soh_cols + ["supply_price","retail_price"]:
    pm[c] = pd.to_numeric(pm[c], errors="coerce")
pm["soh_total"] = pm[soh_cols].fillna(0).sum(axis=1)
pm_unique = pm.drop_duplicates(subset=["sku"], keep="first")

# Build supplier_code → prices lookup as fallback for truncated SKUs
# (when PM SKU is scientific notation and loses precision, we can still
#  match via supplier_code which is always a short alphanumeric string)
pm_by_supplier_code = (
    pm_unique[pm_unique["supplier_code"].notna() & (pm_unique["supplier_code"] != "")]
    .drop_duplicates(subset=["supplier_code"], keep="first")
    [["supplier_code","supply_price","retail_price"]]
    .copy()
)
print(f"   Product Master: {len(pm_unique)} SKUs, {len(pm_by_supplier_code)} with supplier_code")

# ============================================================
# 2. Sales Company
# ============================================================
print("\n[2/9] Loading Sales Company...")

# 先收集所有文件和对应的week_start，按week_start排序后编号
company_file_list = []
for f in sorted(glob.glob(os.path.join(SALES_COMPANY_DIR, "*.csv"))):
    week_start = extract_week_start(os.path.basename(f))
    if week_start is not None:
        company_file_list.append((week_start, f))
company_file_list.sort(key=lambda x: x[0])  # 按日期排序

# 建立 week_start → week_key 映射（按顺序编号）
week_start_to_key = {ws: get_week_key_from_index(i+1) for i, (ws, _) in enumerate(company_file_list)}

company_dfs = []
for week_start, f in company_file_list:
    df = pd.read_csv(f, dtype=str)
    df.rename(columns={
        "Product":"product","SKU":"sku_raw","Supplier Code":"supplier_code",
        "Brand":"brand","Supplier":"supplier","Category":"category","Tag":"tag",
        "Items Sold":"qty_sold","Revenue":"revenue","Cost of Goods Sold":"cogs",
        "Gross Profit":"gross_profit","Margin (%)":"margin_pct","Tax":"tax"
    }, inplace=True)
    df["week_start"] = week_start
    df["week_key"]   = week_start_to_key[week_start]
    df["sku"]        = df["sku_raw"].apply(fix_sku)
    company_dfs.append(df)

fact_company = pd.concat(company_dfs, ignore_index=True)
fact_company["qty_sold"]     = pd.to_numeric(fact_company["qty_sold"],     errors="coerce")
fact_company["revenue"]      = pd.to_numeric(fact_company["revenue"],      errors="coerce")
fact_company["cogs"]         = pd.to_numeric(fact_company["cogs"],         errors="coerce")
fact_company["gross_profit"] = pd.to_numeric(fact_company["gross_profit"], errors="coerce")
fact_company = fact_company[
    fact_company["sku"].notna() & (fact_company["sku"] != "") &
    (fact_company["sku"] != "Error") & fact_company["qty_sold"].notna()
]
fact_company = fact_company[~fact_company["product"].fillna("").str.contains("NZ Post")]
print(f"   Sales Company: {len(fact_company)} rows, {fact_company['week_key'].nunique()} weeks")

# ============================================================
# 3. Sales Store
# ============================================================
print("\n[3/9] Loading Sales Store...")

# 收集store文件，同样按week_start排序，复用company的编号映射
store_file_list = []
for f in sorted(glob.glob(os.path.join(SALES_STORE_DIR, "*.csv"))):
    week_start = extract_week_start(os.path.basename(f))
    if week_start is not None:
        store_file_list.append((week_start, f))
store_file_list.sort(key=lambda x: x[0])

store_dfs = []
for week_start, f in store_file_list:
    df = pd.read_csv(f, dtype=str)
    if "Product" in df.columns:
        df = df[df["Product"] != "Product"]
    df.rename(columns={
        "Product":"product","SKU":"sku_raw","Supplier Code":"supplier_code",
        "Brand":"brand","Supplier":"supplier","Category":"category","Tag":"tag",
        "Outlet":"outlet","Items Sold":"qty_sold","Revenue":"revenue",
        "Cost of Goods Sold":"cogs","Gross Profit":"gross_profit",
        "Margin (%)":"margin_pct","Tax":"tax"
    }, inplace=True)
    df["week_start"] = week_start
    # 优先用company映射，如果store有独立周则也按顺序编号
    df["week_key"]   = week_start_to_key.get(week_start, get_week_key_from_index(
        next((i+1 for i, (ws, _) in enumerate(store_file_list) if ws == week_start), 1)))
    if "sku_raw" in df.columns:
        df["sku"] = df["sku_raw"].apply(fix_sku)
    store_dfs.append(df)

fact_store = pd.concat(store_dfs, ignore_index=True)
fact_store["qty_sold"]     = pd.to_numeric(fact_store["qty_sold"],     errors="coerce")
fact_store["revenue"]      = pd.to_numeric(fact_store["revenue"],      errors="coerce")
fact_store["cogs"]         = pd.to_numeric(fact_store["cogs"],         errors="coerce")
fact_store["gross_profit"] = pd.to_numeric(fact_store["gross_profit"], errors="coerce")
fact_store = fact_store[
    fact_store["sku"].notna() & (fact_store["sku"] != "") &
    (fact_store["sku"] != "Error") & fact_store["qty_sold"].notna()
]
fact_store = fact_store[~fact_store["product"].fillna("").str.contains("NZ Post")]
print(f"   Sales Store: {len(fact_store)} rows, {fact_store['week_key'].nunique()} weeks")

# ============================================================
# 3b. SOH Weekly (Stock On Hand snapshots — one file per week,
#     one row per SKU per outlet, from the "All Inventory" /
#     Product Variant export with Outlet breakdown)
# ============================================================
print("\n[SOH] Loading SOH Weekly...")

soh_file_list = []
for f in sorted(glob.glob(os.path.join(SOH_WEEKLY_DIR, "*.csv"))):
    week_start = extract_week_start(os.path.basename(f))
    if week_start is not None:
        soh_file_list.append((week_start, f))
soh_file_list.sort(key=lambda x: x[0])

soh_dfs = []
SOH_USECOLS = ["SKU Name","SKU","Brand","Supplier","Category","Outlet","Closing Inventory"]
for week_start, f in soh_file_list:
    # Only read the columns we actually use (skip Tag/Supplier Code/Revenue/etc.) —
    # these free-text columns repeated across ~580K rows/week were the main memory cost.
    header = pd.read_csv(f, nrows=0).columns
    usecols = [c for c in SOH_USECOLS if c in header]
    df = pd.read_csv(f, dtype=str, usecols=usecols)
    df.rename(columns={
        "SKU Name":"product","SKU":"sku_raw",
        "Brand":"brand","Supplier":"supplier","Category":"category",
        "Outlet":"outlet","Closing Inventory":"soh_snapshot",
    }, inplace=True)
    df["sku"]          = df["sku_raw"].apply(fix_sku)
    df["soh_snapshot"] = pd.to_numeric(df["soh_snapshot"], errors="coerce")
    # Drop zero-stock rows immediately — they're ~97% of rows (most SKUs are out of
    # stock in most stores) and contribute nothing to any of our sums (0 + x = x),
    # so dropping them here doesn't change any category/supplier total. This is what
    # keeps 29 weeks × ~580K rows/week from ballooning into an 18M-row concat that
    # runs out of memory.
    df = df[df["soh_snapshot"].fillna(0) != 0]
    df["week_start"] = week_start
    # Reuse the SAME week numbering as Sales_Company, matched by week_start,
    # so SOH weeks line up exactly with sales weeks (not renumbered independently).
    df["week_key"] = week_start_to_key.get(week_start)
    for c in ["brand","supplier","category","outlet","week_key"]:
        df[c] = df[c].astype("category")  # memory-efficient for repeated string values
    soh_dfs.append(df)

if soh_dfs:
    fact_soh = pd.concat(soh_dfs, ignore_index=True)
    fact_soh = fact_soh[
        fact_soh["sku"].notna() & (fact_soh["sku"] != "") &
        fact_soh["week_key"].notna() & fact_soh["soh_snapshot"].notna()
    ]
    unmatched = sorted(set(ws for ws, _ in soh_file_list) - set(week_start_to_key.keys()))
    if unmatched:
        print(f"   ⚠️  {len(unmatched)} SOH week(s) had no matching Sales_Company week, skipped: {unmatched}")
    print(f"   SOH Weekly: {len(fact_soh)} rows (zero-stock rows dropped), "
          f"{fact_soh['week_key'].nunique()} weeks ({len(soh_file_list)} files found)")
else:
    fact_soh = pd.DataFrame(columns=[
        "sku","product","brand","supplier","category",
        "outlet","soh_snapshot","week_start","week_key"])
    print("   ⚠️  No SOH Weekly files found in SOH_Weekly folder — skipping SOH history")

# ============================================================
# 4. Week Dimension
# ============================================================
print("\n[4/9] Building Week Dimension...")
dim_week = pd.DataFrame({
    "week_start": [ws for ws, _ in company_file_list],
    "week_key":   [week_start_to_key[ws] for ws, _ in company_file_list],
    "week_number": range(1, len(company_file_list) + 1)
})
MAX_WEEK = dim_week["week_number"].max()
print(f"   Weeks: {len(dim_week)}, Latest: {dim_week['week_key'].iloc[-1]}")

# ============================================================
# 5. Dusty Company
# ============================================================
print("\n[5/9] Loading Dusty Company...")
dusty_co_file = sorted(glob.glob(os.path.join(DUSTY_COMPANY_DIR, "*.csv")),
                        key=os.path.getmtime, reverse=True)[0]
dusty_co = pd.read_csv(dusty_co_file, dtype=str)
col_renames = {
    "Product":"product","SKU":"sku_raw","Supplier Code":"supplier_code",
    "Brand":"brand","Supplier":"supplier","Category":"category","Tag":"tag",
    "Closing Inventory":"soh_company","Created":"created",
    "First Sale":"first_sale","Last Sale":"last_sale",
    "Last received date":"last_received",
    "Inventory Cost":"inventory_cost","Retail Value (Excl. Tax)":"retail_value"
}
if "Items Sold" in dusty_co.columns:
    col_renames["Items Sold"] = "items_sold_lifetime"
dusty_co.rename(columns=col_renames, inplace=True)
dusty_co["sku"] = dusty_co["sku_raw"].apply(fix_sku)
if "items_sold_lifetime" not in dusty_co.columns:
    dusty_co["items_sold_lifetime"] = None
for c in ["soh_company","items_sold_lifetime","inventory_cost","retail_value"]:
    dusty_co[c] = pd.to_numeric(dusty_co[c], errors="coerce")
for c in ["created","first_sale","last_sale","last_received"]:
    dusty_co[c] = dusty_co[c].apply(safe_date)

dusty_co["unit_cost"] = dusty_co.apply(
    lambda r: r["inventory_cost"] / r["soh_company"]
    if pd.notna(r["soh_company"]) and r["soh_company"] != 0 else None, axis=1)
dusty_co["unit_retail"] = dusty_co.apply(
    lambda r: r["retail_value"] / r["soh_company"]
    if pd.notna(r["soh_company"]) and r["soh_company"] != 0 else None, axis=1)
print(f"   Dusty Company: {len(dusty_co)} products")

# ============================================================
# 6. Dusty Store
# ============================================================
print("\n[6/9] Loading Dusty Store...")
dusty_store_dfs = []
for f in glob.glob(os.path.join(DUSTY_STORE_DIR, "*.csv")):
    store_name = os.path.basename(f).replace(".csv","").replace("Dusty_","")
    df = pd.read_csv(f, dtype=str)
    first_col = df.columns[0]
    # Remove repeated header rows regardless of exact text
    df = df[~df[first_col].isin(["Product", "product", first_col])]
    df = df.reset_index(drop=True)
    if "SKU" in df.columns:
        df.rename(columns={
            "Product":"product","SKU":"sku_raw","Supplier Code":"supplier_code",
            "Brand":"brand","Supplier":"supplier","Category":"category","Tag":"tag",
            "Outlet":"outlet_original","Closing Inventory":"soh",
            "Items Sold":"items_sold","Created":"created",
            "First Sale":"first_sale","Last Sale":"last_sale",
            "Inventory Cost":"inventory_cost",
            "Retail Value (Excl. Tax)":"retail_value",
            "Last received date":"last_received"
        }, inplace=True)
    else:
        cols  = list(df.columns)
        names = ["product","sku_raw","supplier_code","brand","supplier",
                 "category","tag","outlet_original","soh","items_sold",
                 "created","first_sale","last_sale","inventory_cost",
                 "retail_value","last_received"]
        df.rename(columns={cols[i]: names[i] for i in range(min(len(cols),len(names)))}, inplace=True)
    df["store"] = store_name
    df["sku"]   = df["sku_raw"].apply(fix_sku)
    for c in ["soh","items_sold","inventory_cost","retail_value"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["created","first_sale","last_sale","last_received"]:
        if c in df.columns:
            df[c] = df[c].apply(safe_date)
    if "soh" in df.columns:
        df = df[df["soh"].notna() & (df["soh"] != 0)]
    df["unit_cost"] = df.apply(
        lambda r: r["inventory_cost"] / r["soh"]
        if pd.notna(r["soh"]) and r["soh"] != 0 else None, axis=1)
    df["unit_retail"] = df.apply(
        lambda r: r["retail_value"] / r["soh"]
        if pd.notna(r["soh"]) and r["soh"] != 0 else None, axis=1)
    dusty_store_dfs.append(df)

dusty_store = pd.concat(dusty_store_dfs, ignore_index=True)
print(f"   Dusty Store total: {len(dusty_store)} rows")

# ============================================================
# 7. Calculations
# ============================================================
print("\n[7/9] Calculating metrics...")

wk_map = dict(zip(dim_week["week_key"], dim_week["week_number"]))
fact_company["week_number"] = fact_company["week_key"].map(wk_map)
fact_store["week_number"]   = fact_store["week_key"].map(wk_map)

fc_wn = fact_company[fact_company["week_number"].notna()].copy()
fc_wn["week_number"] = fc_wn["week_number"].astype(int)
l1_co  = fc_wn[fc_wn["week_number"] >= MAX_WEEK     ].groupby("sku")["qty_sold"].sum().rename("L1W")
l3_co  = fc_wn[fc_wn["week_number"] >= MAX_WEEK - 2 ].groupby("sku")["qty_sold"].sum().rename("L3W")
l6_co  = fc_wn[fc_wn["week_number"] >= MAX_WEEK - 5 ].groupby("sku")["qty_sold"].sum().rename("L6W")
l12_co = fc_wn[fc_wn["week_number"] >= MAX_WEEK - 11].groupby("sku")["qty_sold"].sum().rename("L12W")
sales_agg_co = pd.concat([l1_co,l3_co,l6_co,l12_co], axis=1).fillna(0).reset_index()
sales_agg_co["avg_weekly"] = (sales_agg_co["L6W"] / 6).round(2)

fs_wn = fact_store[fact_store["week_number"].notna()].copy()
fs_wn["week_number"] = fs_wn["week_number"].astype(int)
l1_st  = fs_wn[fs_wn["week_number"] >= MAX_WEEK     ].groupby(["sku","outlet"])["qty_sold"].sum().rename("L1W")
l3_st  = fs_wn[fs_wn["week_number"] >= MAX_WEEK - 2 ].groupby(["sku","outlet"])["qty_sold"].sum().rename("L3W")
l6_st  = fs_wn[fs_wn["week_number"] >= MAX_WEEK - 5 ].groupby(["sku","outlet"])["qty_sold"].sum().rename("L6W")
l12_st = fs_wn[fs_wn["week_number"] >= MAX_WEEK - 11].groupby(["sku","outlet"])["qty_sold"].sum().rename("L12W")
sales_agg_st = pd.concat([l1_st,l3_st,l6_st,l12_st], axis=1).fillna(0).reset_index()
sales_agg_st["avg_weekly"] = (sales_agg_st["L6W"] / 6).round(2)

cy26     = fact_company[fact_company["week_start"].apply(lambda x: x.year if x else 0) == 2026].copy()
cy26_agg = cy26.groupby("sku")["qty_sold"].sum().rename("CY26").reset_index()

fc_2026 = fact_company[fact_company["week_start"].apply(lambda x: x.year if x else 0) == 2026].copy()
fc_2026["wk_label"] = fc_2026["week_key"]
wk_pivot_co = fc_2026.groupby(["sku","wk_label"])["qty_sold"].sum().reset_index().pivot(
    index="sku", columns="wk_label", values="qty_sold")
wk_pivot_co.columns.name = None
wk_pivot_co = wk_pivot_co.reset_index()
wk_cols_co  = sorted([c for c in wk_pivot_co.columns if c != "sku"], reverse=True)

fs_2026 = fact_store[fact_store["week_start"].apply(lambda x: x.year if x else 0) == 2026].copy()
fs_2026["wk_label"] = fs_2026["week_key"]
wk_pivot_st = fs_2026.groupby(["sku","outlet","wk_label"])["qty_sold"].sum().reset_index().pivot_table(
    index=["sku","outlet"], columns="wk_label", values="qty_sold", aggfunc="sum")
wk_pivot_st.columns.name = None
wk_pivot_st = wk_pivot_st.reset_index()
wk_cols_st  = sorted([c for c in wk_pivot_st.columns if c not in ["sku","outlet"]], reverse=True)

latest_wk         = dim_week.iloc[-1]["week_key"]
store_latest_wide = fact_store[fact_store["week_key"] == latest_wk].groupby(
    ["sku","outlet"])["qty_sold"].sum().reset_index().pivot(
    index="sku", columns="outlet", values="qty_sold")
store_latest_wide.columns = [f"store_L1W_{c}" for c in store_latest_wide.columns]
store_latest_wide = store_latest_wide.reset_index()

# ============================================================
# 8. Build Output Company
# ============================================================
print("\n[8/9] Building output tables...")

all_skus = pd.DataFrame({"sku": list(set(dusty_co["sku"].dropna()) | set(pm_unique["sku"].dropna()))})
out_co = all_skus.merge(dusty_co[[
    "sku","product","brand","supplier","supplier_code","category","tag",
    "items_sold_lifetime","created","first_sale","last_sale","last_received",
    "unit_cost","unit_retail"
]], on="sku", how="left")
out_co = out_co.merge(pm_unique[[
    "sku","id","product","supply_price","retail_price","soh_total",
    "soh_Dunedin","soh_Office","soh_Papanui","soh_Queensgate",
    "soh_Riccarton","soh_Richmond","soh_SylviaPark","soh_Terapa",
    "soh_WH1","soh_WarehouseAKL"
]], on="sku", how="left", suffixes=("","_pm"))
out_co["product"] = out_co["product"].fillna(out_co.get("product_pm",""))
if "product_pm" in out_co.columns:
    out_co.drop(columns=["product_pm"], inplace=True)

out_co["inventory_cost"] = (out_co["soh_total"] * out_co["unit_cost"]).round(2)
out_co["retail_value"]   = (out_co["soh_total"] * out_co["unit_retail"]).round(2)
out_co.drop(columns=["unit_cost","unit_retail"], inplace=True)

# Fallback: fill missing supply_price via supplier_code
if "supply_price" in out_co.columns and out_co["supply_price"].isna().any():
    missing_co = out_co["supply_price"].isna() & out_co["supplier_code"].notna()
    if missing_co.any():
        fb_co = out_co[missing_co].merge(
            pm_by_supplier_code, on="supplier_code", how="left", suffixes=("","_fb"))
        out_co.loc[missing_co, "supply_price"] = fb_co["supply_price_fb"].values
        out_co.loc[missing_co, "retail_price"]  = fb_co["retail_price_fb"].values
        filled_co = missing_co.sum() - out_co["supply_price"].isna().sum()
        print(f"   supply_price fallback (company): {filled_co} rows filled")

for col, src in [("days_since_created","created"),("days_since_first_sale","first_sale"),
                  ("days_since_last_sold","last_sale"),("days_since_last_received","last_received")]:
    out_co[col] = out_co[src].apply(days_since)

out_co = out_co.merge(sales_agg_co[["sku","L1W","L3W","L6W","avg_weekly"]], on="sku", how="left")
out_co["WOH"] = out_co.apply(
    lambda r: min(round(r["soh_total"]/r["avg_weekly"],1), 52)
    if pd.notna(r.get("avg_weekly")) and r["avg_weekly"] > 0 else 99, axis=1)
out_co = out_co.merge(cy26_agg, on="sku", how="left")
out_co = out_co.merge(store_latest_wide, on="sku", how="left")

store_l1w_cols = [c for c in out_co.columns if c.startswith("store_L1W_")]
out_co["STORE_TOT"] = out_co[store_l1w_cols].fillna(0).sum(axis=1)

out_co.rename(columns={
    "soh_total":"TOT","soh_Dunedin":"DUN","soh_Office":"Office",
    "soh_Papanui":"PAP","soh_Queensgate":"QG","soh_Riccarton":"RICC",
    "soh_Richmond":"RICH","soh_SylviaPark":"SP","soh_Terapa":"TR",
    "soh_WH1":"WH1","soh_WarehouseAKL":"AKLWH"
}, inplace=True)
out_co.rename(columns={
    "store_L1W_Dunedin":"DUN_L1W","store_L1W_Papanui":"PAP_L1W",
    "store_L1W_Queensgate":"QG_L1W","store_L1W_Riccarton":"RICC_L1W",
    "store_L1W_Richmond":"RICH_L1W","store_L1W_Sylvia Park":"SP_L1W",
    "store_L1W_Te Rapa":"TR_L1W","store_L1W_Office":"Office_L1W"
}, inplace=True)
out_co = out_co.merge(wk_pivot_co, on="sku", how="left")

for c in ["supply_price","retail_price","avg_weekly","WOH",
          "inventory_cost","retail_value"]:
    if c in out_co.columns:
        out_co[c] = pd.to_numeric(out_co[c], errors="coerce").round(2)

# Company Status
def calc_company_status(row):
    soh           = row.get("TOT", 0) or 0
    l6w           = row.get("L6W", 0) or 0
    woh           = row.get("WOH", 999) or 999
    days_sold     = row.get("days_since_last_sold", 999) or 999
    days_created  = row.get("days_since_created", 999) or 999
    days_first    = row.get("days_since_first_sale", 999) or 999
    days_received = row.get("days_since_last_received", 999) or 999
    if soh <= 0:
        return "No Stock"
    elif days_created <= 28 or days_first <= 28:
        return "New"
    elif days_received <= 14:
        return "New (Replenished)"
    elif l6w == 0 and days_sold >= 42:
        return "Dead"
    elif l6w >= 6 and woh <= 10:
        return "Good"
    else:
        return "Slow"

def calc_company_action(row):
    cs  = row.get("Company_Status", "")
    woh = row.get("WOH", 999) or 999
    if cs == "Dead":
        return "STOP/CLR"
    elif cs in ["Good", "New", "New (Replenished)"]:
        if woh <= 4:
            return "URGENT RESTOCK"
        elif woh <= 8:
            return "RESTOCK"
        else:
            return "MONITOR"
    else:
        return "MONITOR"

out_co["Company_Status"] = out_co.apply(calc_company_status, axis=1)
out_co["Action"]         = out_co.apply(calc_company_action, axis=1)

fixed_cols_co = [
    "id","sku","product","category","tag",
    "supply_price","retail_price","inventory_cost","retail_value",
    "brand","supplier","supplier_code",
    "DUN","PAP","QG","RICC","RICH","SP","TR","Office","WH1","AKLWH","TOT",
    "created","first_sale","last_sale","last_received",
    "days_since_created","days_since_first_sale",
    "days_since_last_sold","days_since_last_received",
    "DUN_L1W","PAP_L1W","QG_L1W","RICC_L1W","RICH_L1W",
    "SP_L1W","TR_L1W","Office_L1W","STORE_TOT",
    "items_sold_lifetime","CY26","avg_weekly","WOH","L6W","L3W","L1W",
    "Company_Status","Action"
]
existing_fixed = [c for c in fixed_cols_co if c in out_co.columns]
wk_cols_sorted = sorted([c for c in wk_cols_co if c in out_co.columns], reverse=True)
extra_cols     = [c for c in out_co.columns if c not in fixed_cols_co and c not in wk_cols_sorted]
out_co         = out_co[existing_fixed + wk_cols_sorted + extra_cols]
print(f"   Output Company: {len(out_co)} rows")

# ============================================================
# Build Output Store
# ============================================================
out_store = dusty_store[[
    "store","sku","product","brand","supplier","supplier_code",
    "category","tag","soh","items_sold",
    "created","first_sale","last_sale","last_received",
    "unit_cost","unit_retail"
]].copy()

outlet_map = {
    "Dunedin":"Dunedin","Papanui":"Papanui","Queensgate":"Queensgate",
    "Riccarton":"Riccarton","Richmond":"Richmond","SylviaPark":"Sylvia Park",
    "Terapa":"Te Rapa","Office":"Office","WarehouseAKL":"Warehouse AKL","WH1":"WH1"
}
out_store["outlet"] = out_store["store"].map(outlet_map).fillna(out_store["store"])
out_store = out_store.merge(
    pm_unique[["sku","supply_price","retail_price"]],
    on="sku", how="left", suffixes=("_dusty",""))

# Fallback: for rows where supply_price is still NaN (truncated SKU in PM),
# try to match via supplier_code
if "supply_price" in out_store.columns and out_store["supply_price"].isna().any():
    missing_mask = out_store["supply_price"].isna() & out_store["supplier_code"].notna()
    if missing_mask.any():
        fallback = out_store[missing_mask].merge(
            pm_by_supplier_code,
            on="supplier_code", how="left", suffixes=("","_fb")
        )
        out_store.loc[missing_mask, "supply_price"] = fallback["supply_price_fb"].values
        out_store.loc[missing_mask, "retail_price"]  = fallback["retail_price_fb"].values
        filled = missing_mask.sum() - out_store["supply_price"].isna().sum()
        print(f"   supply_price fallback via supplier_code: {filled} rows filled")
out_store = out_store.merge(
    sales_agg_st[["sku","outlet","L1W","L3W","L6W","avg_weekly"]],
    on=["sku","outlet"], how="left")
out_store["WOH_store"] = out_store.apply(
    lambda r: min(round(r["soh"]/r["avg_weekly"],1), 52)
    if pd.notna(r.get("avg_weekly")) and r["avg_weekly"] > 0 else 99, axis=1)
out_store["inventory_cost"] = (out_store["soh"] * out_store["unit_cost"]).round(2)
out_store["retail_value"]   = (out_store["soh"] * out_store["unit_retail"]).round(2)
out_store.drop(columns=["unit_cost","unit_retail"], inplace=True)

for col, src in [("days_since_last_sold","last_sale"),
                  ("days_since_last_received","last_received"),
                  ("days_since_created","created"),
                  ("days_since_first_sale","first_sale")]:
    out_store[col] = out_store[src].apply(days_since)

out_store = out_store.merge(
    out_co[["sku","TOT","L1W","L3W","L6W","avg_weekly","WOH",
            "items_sold_lifetime","Company_Status"]].rename(columns={
        "TOT":"SOH_co","L1W":"L1W_co","L3W":"L3W_co","L6W":"L6W_co",
        "avg_weekly":"avg_weekly_co","WOH":"WOH_co",
        "items_sold_lifetime":"items_sold_lifetime_co"
    }), on="sku", how="left")

out_store = out_store.merge(cy26_agg, on="sku", how="left")
out_store = out_store.merge(wk_pivot_st, on=["sku","outlet"], how="left")

for c in ["soh","SOH_co","L1W","L1W_co","L3W","L3W_co","L6W","L6W_co",
          "avg_weekly","avg_weekly_co","WOH_store","WOH_co","CY26",
          "inventory_cost","retail_value"]:
    if c in out_store.columns:
        out_store[c] = pd.to_numeric(out_store[c], errors="coerce").round(2)

def calc_store_status(row):
    soh           = row.get("soh", 0) or 0
    days_created  = row.get("days_since_created", 999) or 999
    days_first    = row.get("days_since_first_sale", 999) or 999
    days_sold     = row.get("days_since_last_sold", 999) or 999
    days_received = row.get("days_since_last_received", 999) or 999
    l6w           = row.get("L6W", 0) or 0
    woh           = row.get("WOH_store", 999) or 999
    if soh <= 0:
        return "No Stock"
    elif days_created <= 28 or days_first <= 28:
        return "New"
    elif days_received <= 14:
        return "New (Replenished)"
    elif l6w == 0 and days_sold >= 42:
        return "Dead"
    elif l6w >= 6 and woh <= 10:
        return "Good"
    else:
        return "Slow"

def calc_store_action(row):
    ss     = row.get("Store_Status", "")
    cs     = row.get("Company_Status", "")
    woh_st = row.get("WOH_store", 999) or 999
    if ss == "Dead" and cs == "Dead":
        return "STOP/CLR"
    elif ss == "Dead" and cs in ["Good", "New", "New (Replenished)"]:
        return "TRANSFER_OUT"
    elif ss == "Dead":
        return "TRANSFER/CLR"
    elif ss in ["Good", "New", "New (Replenished)"]:
        if woh_st <= 4:
            return "URGENT RESTOCK" if cs == "Good" else "CHECK TRANSFER_IN"
        return "RESTOCK" if woh_st <= 8 else "MONITOR"
    elif ss != "Good" and cs == "Good":
        return "TRANSFER_IN"
    else:
        return "MONITOR"

out_store["Store_Status"] = out_store.apply(calc_store_status, axis=1)
out_store["Action"]       = out_store.apply(calc_store_action, axis=1)
out_store = out_store.sort_values("sku").reset_index(drop=True)

fixed_cols_st = [
    "store","sku","product","category","tag","supplier","supplier_code","brand",
    "supply_price","retail_price","inventory_cost","retail_value",
    "soh","SOH_co",
    "created","first_sale","last_sale","last_received",
    "days_since_last_sold","days_since_last_received",
    "L1W","L1W_co","L3W","L3W_co","L6W","L6W_co",
    "avg_weekly","avg_weekly_co","WOH_store","WOH_co",
    "CY26","items_sold","items_sold_lifetime_co",
    "Store_Status","Company_Status","Action"
]
existing_fixed_st = [c for c in fixed_cols_st if c in out_store.columns]
wk_cols_st_sorted = sorted([c for c in wk_cols_st if c in out_store.columns], reverse=True)
extra_cols_st     = [c for c in out_store.columns if c not in fixed_cols_st and c not in wk_cols_st_sorted]
out_store         = out_store[existing_fixed_st + wk_cols_st_sorted + extra_cols_st]
print(f"   Output Store: {len(out_store)} rows")

# ============================================================
# 9. Weekly Reports
# ============================================================
# ============================================================
# 9. Weekly Reports
# ============================================================
print("\n[9/9] Generating weekly reports...")

fc_2026_rep = fact_company[
    fact_company["week_start"].apply(lambda x: x.year if x else 0) == 2026
].copy()
fs_2026_rep = fact_store[
    fact_store["week_start"].apply(lambda x: x.year if x else 0) == 2026
].copy()

for c in ["qty_sold","revenue","cogs","gross_profit"]:
    fc_2026_rep[c] = pd.to_numeric(fc_2026_rep[c], errors="coerce").fillna(0)
    fs_2026_rep[c] = pd.to_numeric(fs_2026_rep[c], errors="coerce").fillna(0)

# 按文件顺序排序的week列表（W01→W18）
all_weeks = [wk for wk in dim_week["week_key"].tolist() if wk in fc_2026_rep["week_key"].unique()]

# 只保留8家门店，去掉WarehouseAKL和WH1
VALID_STORES = ["Dunedin", "Papanui", "Queensgate", "Riccarton",
                "Richmond", "Sylvia Park", "Te Rapa", "Office"]
fs_2026_rep = fs_2026_rep[fs_2026_rep["outlet"].isin(VALID_STORES)].copy()

# 提取Brand/IP（产品名 - 前的部分）
def extract_brand(product_name):
    if pd.isna(product_name):
        return "Unknown"
    parts = str(product_name).split(" - ")
    return parts[0].strip() if len(parts) > 1 else str(product_name).strip()

fc_2026_rep["brand_ip"] = fc_2026_rep["product"].apply(extract_brand)
fs_2026_rep["brand_ip"] = fs_2026_rep["product"].apply(extract_brand)

# 提取主Tag（第一个tag）
def extract_main_tag(tag_str):
    if pd.isna(tag_str):
        return "No Tag"
    tags = str(tag_str).split(",")
    return tags[0].strip() if tags else "No Tag"

fc_2026_rep["main_tag"] = fc_2026_rep["tag"].apply(extract_main_tag)
fs_2026_rep["main_tag"] = fs_2026_rep["tag"].apply(extract_main_tag)

def build_pivot_report(df, group_col, weeks):
    result = {}
    for metric, val_col, agg in [
        ("REV",    "revenue",      "sum"),
        ("GP",     "gross_profit", "sum"),
        ("SALES_U","qty_sold",     "sum"),
        ("SKU",    "sku",          "nunique"),
    ]:
        if agg == "sum":
            pivot = df.groupby([group_col,"week_key"])[val_col].sum().reset_index()
        else:
            pivot = df.groupby([group_col,"week_key"])[val_col].nunique().reset_index()
        pivot = pivot.pivot(index=group_col, columns="week_key", values=val_col)
        pivot = pivot.reindex(columns=weeks).fillna(0)
        pivot.columns.name = None
        pivot = pivot.reset_index()
        pivot = pivot.sort_values(group_col)
        result[metric] = pivot

    margin = result["REV"].copy()
    for wk in weeks:
        rev = result["REV"][wk]
        gp  = result["GP"][wk]
        margin[wk] = (gp / rev.replace(0, float("nan"))).fillna(0).round(4)
    result["MARGIN"] = margin
    return result

def build_pivot_with_pct(df, group_col, weeks, pct_base_df=None):
    """建立带占比的pivot，pct_base_df是计算占比的基准（None则用自身）"""
    result = build_pivot_report(df, group_col, weeks)

    # 计算REV占比
    if pct_base_df is None:
        base_rev = df.groupby("week_key")["revenue"].sum()
    else:
        base_rev = pct_base_df.groupby("week_key")["revenue"].sum()

    pct_pivot = result["REV"].copy()
    for wk in weeks:
        total = base_rev.get(wk, 0)
        if total > 0:
            pct_pivot[wk] = (result["REV"][wk] / total).round(4)
        else:
            pct_pivot[wk] = 0
    result["REV_PCT"] = pct_pivot
    return result

# Category报表
cat_company = build_pivot_report(fc_2026_rep, "category", all_weeks)
store_names_rep = sorted(fs_2026_rep["outlet"].dropna().unique().tolist())
cat_stores = {}
for s in store_names_rep:
    sdf = fs_2026_rep[fs_2026_rep["outlet"] == s]
    cat_stores[s] = build_pivot_report(sdf, "category", all_weeks)

# Supplier报表
sup_company = build_pivot_report(fc_2026_rep, "supplier", all_weeks)

# Brand/IP报表 — 公司整体
brand_company = build_pivot_with_pct(fc_2026_rep, "brand_ip", all_weeks)

# Brand/IP报表 — 按Category细分（公司）
brand_by_cat_company = {}
for cat in sorted(fc_2026_rep["category"].dropna().unique()):
    cdf = fc_2026_rep[fc_2026_rep["category"] == cat]
    if len(cdf) > 0:
        brand_by_cat_company[cat] = build_pivot_with_pct(cdf, "brand_ip", all_weeks, pct_base_df=cdf)

# Brand/IP报表 — 按Store细分
brand_stores = {}
for s in store_names_rep:
    sdf = fs_2026_rep[fs_2026_rep["outlet"] == s]
    brand_stores[s] = build_pivot_with_pct(sdf, "brand_ip", all_weeks)

# Brand/IP报表 — 按Store+Category细分
brand_by_cat_stores = {}
for s in store_names_rep:
    sdf = fs_2026_rep[fs_2026_rep["outlet"] == s]
    brand_by_cat_stores[s] = {}
    for cat in sorted(sdf["category"].dropna().unique()):
        cdf = sdf[sdf["category"] == cat]
        if len(cdf) > 0:
            brand_by_cat_stores[s][cat] = build_pivot_with_pct(cdf, "brand_ip", all_weeks, pct_base_df=cdf)

# Tag报表 — 公司整体
tag_company = build_pivot_with_pct(fc_2026_rep, "main_tag", all_weeks)

# Tag报表 — 按Category细分（公司）
tag_by_cat_company = {}
for cat in sorted(fc_2026_rep["category"].dropna().unique()):
    cdf = fc_2026_rep[fc_2026_rep["category"] == cat]
    if len(cdf) > 0:
        tag_by_cat_company[cat] = build_pivot_with_pct(cdf, "main_tag", all_weeks, pct_base_df=cdf)

# Tag报表 — 按Store细分
tag_stores = {}
for s in store_names_rep:
    sdf = fs_2026_rep[fs_2026_rep["outlet"] == s]
    tag_stores[s] = build_pivot_with_pct(sdf, "main_tag", all_weeks)

# ============================================================
# SOH History & Category WOH Trend (from SOH Weekly snapshots)
# ============================================================
print("\n[SOH] Building SOH history & WOH trend reports...")

def build_soh_pivot(df, group_col, weeks):
    """Same shape as build_pivot_report's REV/GP tables, but for a single
    snapshot metric (SOH) rather than a summed-per-week flow metric."""
    pivot = df.groupby([group_col, "week_key"])["soh_snapshot"].sum().reset_index()
    pivot = pivot.pivot(index=group_col, columns="week_key", values="soh_snapshot")
    pivot = pivot.reindex(columns=weeks).fillna(0)
    pivot.columns.name = None
    return pivot.reset_index().sort_values(group_col)

if len(fact_soh) > 0:
    # Company-level SOH per SKU per week: sum across ALL outlets. Grouping
    # only by sku (not brand/category, which have NaN gaps) avoids pandas
    # silently dropping rows with a missing group key — same fix as the
    # earlier Excel reconciliation issue.
    soh_sku_week = fact_soh.groupby(["sku", "week_key"], as_index=False)["soh_snapshot"].sum()
    sku_lookup   = fact_soh.drop_duplicates(subset="sku")[["sku", "category", "supplier", "brand"]]
    soh_sku_week = soh_sku_week.merge(sku_lookup, on="sku", how="left")

    soh_by_category_company = build_soh_pivot(soh_sku_week, "category", all_weeks)
    soh_by_supplier_company = build_soh_pivot(soh_sku_week, "supplier", all_weeks)

    # WOH by category, per week = that week's category SOH / that week's category
    # units sold (so you can see whether cover is improving or worsening over time,
    # not just the current snapshot).
    sales_by_cat_week = fc_2026_rep.groupby(["category", "week_key"])["qty_sold"].sum().reset_index()
    sales_pivot_cat = sales_by_cat_week.pivot(
        index="category", columns="week_key", values="qty_sold"
    ).reindex(columns=all_weeks).fillna(0)
    soh_cat_indexed = soh_by_category_company.set_index("category")[all_weeks]
    sales_cat_aligned = sales_pivot_cat.reindex(soh_cat_indexed.index).fillna(0)
    woh_by_category_company = (soh_cat_indexed / sales_cat_aligned.replace(0, float("nan")))
    woh_by_category_company = woh_by_category_company.fillna(99).round(1).reset_index()

    # Store-level SOH by category, one pivot per store (mirrors cat_stores structure)
    soh_stores_by_category = {}
    for s in store_names_rep:
        sdf = fact_soh[fact_soh["outlet"] == s]
        if len(sdf) > 0:
            soh_stores_by_category[s] = build_soh_pivot(sdf, "category", all_weeks)

    soh_history = {
        "weeks":                   [wk for wk in all_weeks if wk in fact_soh["week_key"].unique()],
        "soh_by_category_company": soh_by_category_company,
        "soh_by_supplier_company": soh_by_supplier_company,
        "woh_by_category_company": woh_by_category_company,
        "soh_stores_by_category":  soh_stores_by_category,
    }
    print(f"   SOH history: {soh_sku_week['week_key'].nunique()} weeks, "
          f"{len(soh_by_category_company)} categories, {len(soh_by_supplier_company)} suppliers")
else:
    soh_history = {}
    print("   Skipped — no SOH data loaded")

# 保存
weekly_data = {
    "weeks":               all_weeks,
    "cat_company":         cat_company,
    "cat_stores":          cat_stores,
    "sup_company":         sup_company,
    "brand_company":       brand_company,
    "brand_by_cat_company":brand_by_cat_company,
    "brand_stores":        brand_stores,
    "brand_by_cat_stores": brand_by_cat_stores,
    "tag_company":         tag_company,
    "tag_by_cat_company":  tag_by_cat_company,
    "tag_stores":          tag_stores,
    "soh_history":         soh_history,
}
with open(os.path.join(OUTPUT_DIR, "weekly_reports.pkl"), "wb") as f:
    pickle.dump(weekly_data, f)

print(f"   Weekly reports: {len(all_weeks)} weeks, {len(store_names_rep)} stores")
# ============================================================
# Save
# ============================================================
print("\nSaving data...")
out_co.to_parquet(os.path.join(OUTPUT_DIR, "output_company.parquet"), index=False)
out_store.to_parquet(os.path.join(OUTPUT_DIR, "output_store.parquet"), index=False)
pd.DataFrame({"store": out_store["store"].dropna().unique().tolist()}).to_parquet(
    os.path.join(OUTPUT_DIR, "stores.parquet"), index=False)

print("\n" + "=" * 60)
print("Data processing complete!")
print(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("=" * 60)