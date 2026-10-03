"""Automated data cleaning + reporting pipeline.
Usage:  python clean_and_report.py [--input data/raw_sales.csv] [--out outputs]
Produces: cleaned CSV, Excel report (summary, cleaning log, pivots, charts), HTML report, PNG charts, run log."""
import argparse, base64, io, datetime as dt
import numpy as np, pandas as pd, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt, seaborn as sns
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment
sns.set_theme(style="whitegrid")

ap = argparse.ArgumentParser()
ap.add_argument("--input", default="data/raw_sales.csv"); ap.add_argument("--out", default="outputs")
a = ap.parse_args()
log = []                                                       # (step, issue, rows/cells affected, action)
def note(step, issue, n, action): log.append({"Step": step, "Issue found": issue, "Affected": int(n), "Action taken": action}); print(f"[{step}] {issue}: {int(n)} -> {action}")

# ---------- 1. LOAD ----------
raw = pd.read_csv(a.input, dtype=str); df = raw.copy()
df.columns = df.columns.str.strip()
rows_in = len(df); miss_before = df.isna().sum()
print(f"Loaded {rows_in} rows, {df.shape[1]} columns from {a.input}")

# ---------- 2. TEXT CLEANUP ----------
txt = ["Order_ID","Customer_Name","Email","Product","Category","Region","Payment_Method","Status"]
ws = sum((df[c].dropna() != df[c].dropna().str.strip()).sum() for c in txt)
for c in txt: df[c] = df[c].str.strip()
note("Text", "Leading/trailing spaces", ws, "trimmed")
df["Customer_Name"] = df.Customer_Name.str.title()
df["Email"] = df.Email.str.lower()
df["Product"] = df.Product.str.title()

# ---------- 3. STANDARDISE CATEGORIES ----------
maps = {
 "Region": {"n.":"North","s.":"South","e.":"East","w.":"West"},
 "Payment_Method": {"cc":"Credit Card","credit-card":"Credit Card","credit card":"Credit Card","debit card":"Debit Card","upi":"UPI","cash":"Cash"},
 "Status": {"done":"Completed","completed":"Completed","returned":"Returned","cancelled":"Cancelled"},
}
for c, m in maps.items():
    before = df[c].copy(); low = df[c].str.lower()
    df[c] = low.map(m).fillna(low.str.title()).where(df[c].notna())
    note("Standardise", f"Inconsistent labels in {c}", (before.notna() & (before != df[c])).sum(), "mapped to one standard spelling")
cat_map = df.dropna(subset=["Category"]).groupby("Product").Category.agg(lambda s: s.mode().iloc[0])
df["Category"] = df.Product.map(cat_map).fillna(df.Category)

# ---------- 4. TYPES ----------
dd = pd.to_datetime(df.Order_Date, format="mixed", errors="coerce", dayfirst=False)
note("Types", "Dates in mixed formats (ISO, dd-Mon-yyyy, mm/dd/yyyy)", df.Order_Date.notna().sum(), "parsed to one date type")
df["Order_Date"] = dd
price_text = df.Unit_Price.str.contains(r"[\$,]", na=False).sum()
df["Unit_Price"] = pd.to_numeric(df.Unit_Price.str.replace(r"[\$,]", "", regex=True), errors="coerce")
note("Types", "Prices stored as text with $ and commas", price_text, "converted to numbers")
df["Quantity"] = pd.to_numeric(df.Quantity, errors="coerce")

# ---------- 5. DUPLICATES ----------
n = df.duplicated().sum(); df = df.drop_duplicates(); note("Duplicates", "Exact duplicate rows", n, "removed")
n = df.duplicated("Order_ID").sum(); df = df.sort_values("Order_Date").drop_duplicates("Order_ID", keep="last")
note("Duplicates", "Repeated Order_ID (updated records)", n, "kept latest version only")

# ---------- 6. INVALID VALUES ----------
bad_q = (df.Quantity <= 0).sum(); df.loc[df.Quantity <= 0, "Quantity"] = np.nan
note("Validation", "Zero/negative quantities", bad_q, "set to missing, then imputed")
bad_e = (df.Email.notna() & ~df.Email.str.match(r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$", na=False)).sum()
df.loc[df.Email.notna() & ~df.Email.str.match(r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$", na=False), "Email"] = np.nan
note("Validation", "Invalid email addresses", bad_e, "set to missing")
n = df.Order_Date.isna().sum(); df = df.dropna(subset=["Order_Date"]); note("Missing", "Rows without an order date", n, "dropped (cannot be placed in time)")

# ---------- 7. MISSING VALUES ----------
for col in ["Quantity","Unit_Price"]:
    n = df[col].isna().sum(); df[col] = df[col].fillna(df.groupby("Product")[col].transform("median"))
    note("Missing", f"Missing {col}", n, "filled with the median for the same product")
for col in ["Customer_Name","Region"]:
    n = df[col].isna().sum(); df[col] = df[col].fillna("Unknown"); note("Missing", f"Missing {col}", n, "filled with 'Unknown'")
n = df.Email.isna().sum(); note("Missing", "Missing/invalid Email", n, "left blank (not needed for revenue)")

# ---------- 8. FEATURES & OUTPUT ----------
df["Revenue"] = (df.Quantity * df.Unit_Price).round(2); df["Month"] = df.Order_Date.dt.to_period("M").astype(str)
df = df.sort_values("Order_Date").reset_index(drop=True)
df.to_csv(f"{a.out}/clean_sales.csv", index=False)
rows_out = len(df); miss_after = df[raw.columns].isna().sum()
logdf = pd.DataFrame(log)

# ---------- 9. ANALYSIS ----------
ok = df[df.Status == "Completed"]
kpi = {"Raw rows": rows_in, "Clean rows": rows_out, "Rows removed": rows_in-rows_out,
       "Cells fixed or filled": int(logdf.Affected.sum()), "Completed orders": len(ok),
       "Total revenue (completed)": round(ok.Revenue.sum(), 2), "Average order value": round(ok.Revenue.mean(), 2),
       "Return rate %": round((df.Status == "Returned").mean()*100, 1), "Cancel rate %": round((df.Status == "Cancelled").mean()*100, 1)}
monthly = ok.groupby("Month").agg(Orders=("Order_ID","count"), Revenue=("Revenue","sum")).round(0).reset_index()
by_region = ok.groupby("Region").Revenue.sum().round(0).sort_values(ascending=False).reset_index()
by_cat = ok.groupby("Category").Revenue.sum().round(0).sort_values(ascending=False).reset_index()
top_prod = ok.groupby("Product").agg(Units=("Quantity","sum"), Revenue=("Revenue","sum")).round(0).sort_values("Revenue", ascending=False).reset_index()
by_pay = ok.groupby("Payment_Method").Revenue.sum().round(0).sort_values(ascending=False).reset_index()

# ---------- 10. CHARTS ----------
C = {}
def save(name): plt.tight_layout(); plt.savefig(f"{a.out}/{name}.png", dpi=120); plt.close(); C[name] = f"{a.out}/{name}.png"
mb = miss_before.reindex(raw.columns).fillna(0); ma = miss_after.reindex(raw.columns).fillna(0)
q = pd.DataFrame({"Before": mb, "After": ma}); q = q[(q.Before > 0) | (q.After > 0)]
q.plot(kind="bar", figsize=(8,4), color=["#d9534f","#5cb85c"]); plt.title("Missing values per column: before vs after"); plt.xticks(rotation=30, ha="right"); save("01_data_quality")
plt.figure(figsize=(9,4)); plt.plot(monthly.Month, monthly.Revenue, marker="o"); plt.xticks(rotation=60, ha="right"); plt.title("Monthly revenue (completed orders)"); plt.ylabel("Revenue"); save("02_monthly_revenue")
plt.figure(figsize=(6,4)); sns.barplot(data=by_region, x="Region", y="Revenue", color="#3b82f6"); plt.title("Revenue by region"); save("03_revenue_by_region")
plt.figure(figsize=(6,4)); sns.barplot(data=top_prod, y="Product", x="Revenue", color="#10b981"); plt.title("Revenue by product"); save("04_revenue_by_product")
plt.figure(figsize=(5,4)); plt.pie(by_cat.Revenue, labels=by_cat.Category, autopct="%1.0f%%", startangle=90); plt.title("Revenue share by category"); save("05_category_share")

# ---------- 11. EXCEL REPORT ----------
xl = f"{a.out}/sales_report.xlsx"
with pd.ExcelWriter(xl, engine="openpyxl") as w:
    pd.DataFrame({"Metric": kpi.keys(), "Value": kpi.values()}).to_excel(w, sheet_name="Summary", index=False)
    logdf.to_excel(w, sheet_name="Cleaning Log", index=False)
    monthly.to_excel(w, sheet_name="Monthly Revenue", index=False); by_region.to_excel(w, sheet_name="By Region", index=False)
    by_cat.to_excel(w, sheet_name="By Category", index=False); top_prod.to_excel(w, sheet_name="By Product", index=False)
    by_pay.to_excel(w, sheet_name="By Payment", index=False); df.to_excel(w, sheet_name="Clean Data", index=False)
wb = load_workbook(xl)
for s in wb.worksheets:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E79"); c.alignment = Alignment(horizontal="center")
    for col in s.columns: s.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or "")) for c in col[:60]) + 3, 55)
    s.freeze_panes = "A2"
ch = wb.create_sheet("Charts", 1)
for i, (k, p) in enumerate(C.items()): img = XLImage(p); img.width, img.height = 520, 300; ch.add_image(img, f"{'A' if i%2==0 else 'L'}{1+(i//2)*17}")
wb.save(xl)

# ---------- 12. HTML REPORT ----------
def b64(p): return base64.b64encode(open(p, "rb").read()).decode()
now = dt.datetime.now().strftime("%d %b %Y %H:%M")
html = f"""<html><head><meta charset="utf-8"><title>Sales Data Report</title><style>body{{font-family:Segoe UI,Arial;max-width:1000px;margin:auto;padding:20px;color:#222}}
h1{{color:#1f4e79}}table{{border-collapse:collapse;width:100%;margin:10px 0}}th{{background:#1f4e79;color:#fff}}td,th{{border:1px solid #ddd;padding:6px;font-size:14px}}
.k{{display:flex;flex-wrap:wrap;gap:10px}}.k div{{background:#eef4fb;padding:12px 16px;border-radius:8px}}.k b{{display:block;font-size:20px}}img{{max-width:48%;margin:4px}}</style></head><body>
<h1>Automated Sales Data Report</h1><p>Generated {now} from <code>{a.input}</code></p>
<div class="k">{''.join(f'<div>{k}<b>{v:,}</b></div>' for k,v in kpi.items())}</div>
<h2>What was cleaned</h2>{logdf.to_html(index=False)}
<h2>Visual summary</h2>{''.join(f'<img src="data:image/png;base64,{b64(p)}">' for p in C.values())}
<h2>Top products</h2>{top_prod.to_html(index=False)}</body></html>"""
open(f"{a.out}/sales_report.html", "w", encoding="utf-8").write(html)
open(f"{a.out}/run_log.txt", "a").write(f"{now} | in={rows_in} out={rows_out} | revenue={kpi['Total revenue (completed)']:,}\n")
print("\nDONE ->", a.out); print(pd.Series(kpi).to_string())
