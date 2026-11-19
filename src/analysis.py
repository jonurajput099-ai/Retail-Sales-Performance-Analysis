"""
Retail Sales Performance Analysis
Project 1 - Python, NumPy & Pandas
"""
import numpy as np
import pandas as pd

pd.set_option("display.width", 120)
pd.set_option("display.max_columns", 20)

DATA = r"C:\Users\Lenovo\Documents\Retail-Sales-Analysis\data\retail_sales.csv"
OUT = r"C:\Users\Lenovo\Documents\Retail-Sales-Analysis\src\results.txt"

lines = []
def show(title, obj):
    lines.append("\n" + "=" * 72)
    lines.append(title)
    lines.append("=" * 72)
    text = obj if isinstance(obj, str) else obj.to_string()
    lines.append(text)
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)
    print(text)

# ---------- 1. Load & inspect ----------
df = pd.read_csv(DATA)
show("1. DATASET SHAPE", f"rows = {len(df)}, columns = {len(df.columns)}")
show("1b. COLUMNS & DTYPES", pd.DataFrame({"column": df.columns, "dtype": df.dtypes.astype(str)}))
show("1c. FIRST 5 ROWS", df.head())

# ---------- 2. Missing values & duplicates ----------
na = df.isna().sum()
show("2. MISSING VALUES", na[na > 0].to_string() if (na > 0).any() else "No missing values")
dups = df[df.duplicated(subset=["Order_ID"], keep=False)].sort_values("Order_ID")
show("2b. DUPLICATE ORDER_ID ROWS", dups if len(dups) else "No duplicates")
show("2c. FULL-ROW DUPLICATES", f"{df.duplicated().sum()}")

# ---------- 3. Clean City (mode per Category, fallback global mode) ----------
global_mode = df["City"].mode()[0]
cat_mode = df.groupby("Category")["City"].agg(lambda s: s.mode().iat[0] if len(s.mode()) else global_mode)
mask = df["City"].isna()
df.loc[mask, "City"] = df.loc[mask, "Category"].map(cat_mode).fillna(global_mode)
show("3. CITY CLEANING", f"filled {mask.sum()} missing cities using per-category mode")
show("3b. CITY VALUE COUNTS", df["City"].value_counts().to_string())

# ---------- Remove duplicates ----------
df = df.drop_duplicates(subset=["Order_ID"], keep="first").copy()
show("3c. AFTER DEDUP", f"rows now = {len(df)}")

# ---------- 4. Date handling ----------
df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")
df["Month"] = df["Order_Date"].dt.month
df["Month_Name"] = df["Order_Date"].dt.month_name()
df["Weekday"] = df["Order_Date"].dt.day_name()
show("4. DATE RANGE", f"{df['Order_Date'].min().date()} to {df['Order_Date'].max().date()}")

# ---------- 5. Revenue verification ----------
df["Calc_Revenue"] = df["Quantity"] * df["Unit_Price"]
diff = (df["Calc_Revenue"] - df["Revenue"]).abs()
show("5. REVENUE CHECK",
     f"max difference = {diff.max():.4f}\nrows differing > 0.01 = {(diff > 0.01).sum()}\nsupplied Revenue is consistent with Quantity x Unit_Price")
df["Revenue"] = df["Revenue"].fillna(df["Calc_Revenue"])

# ---------- 6. Totals ----------
show("6. BUSINESS TOTALS",
     f"Total revenue   : {df['Revenue'].sum():,.2f}\n"
     f"Total quantity  : {df['Quantity'].sum():,}\n"
     f"Orders          : {df['Order_ID'].nunique():,}\n"
     f"Average order value: {df['Revenue'].mean():,.2f}")

# ---------- 7. Top products by revenue ----------
show("7. TOP 10 PRODUCTS BY REVENUE",
     df.groupby("Product")["Revenue"].sum().sort_values(ascending=False).head(10).round(2).to_frame())

# ---------- 8. Cities by revenue ----------
city = (df.groupby("City")
          .agg(Revenue=("Revenue", "sum"), Orders=("Order_ID", "nunique"))
          .sort_values("Revenue", ascending=False).round(2))
city["Avg_Order_Value"] = (city["Revenue"] / city["Orders"]).round(2)
show("8. CITY PERFORMANCE (revenue desc)", city)

# ---------- 9. Monthly revenue ----------
monthly = (df.groupby(["Month", "Month_Name"])["Revenue"]
             .sum().reset_index().sort_values("Month"))
monthly["Revenue"] = monthly["Revenue"].round(2)
monthly["Share_%"] = (monthly["Revenue"] / monthly["Revenue"].sum() * 100).round(1)
show("9. MONTHLY REVENUE", monthly)
best = monthly.loc[monthly["Revenue"].idxmax()]
show("9b. BEST MONTH", f"{best['Month_Name']} with revenue {best['Revenue']:,.2f}")

# ---------- 10. Top customers ----------
cust = (df.groupby("Customer_ID")
          .agg(Revenue=("Revenue", "sum"), Orders=("Order_ID", "nunique"))
          .sort_values("Revenue", ascending=False))
cust["Avg_Order_Value"] = (cust["Revenue"] / cust["Orders"]).round(2)
show("10. TOP 10 CUSTOMERS", cust.head(10).round(2))

# ---------- 11. Payment methods ----------
pay = (df.groupby("Payment_Mode")
         .agg(Orders=("Order_ID", "nunique"), Revenue=("Revenue", "sum"))
         .sort_values("Revenue", ascending=False))
pay["Share_%"] = (pay["Revenue"] / pay["Revenue"].sum() * 100).round(1)
pay["Revenue"] = pay["Revenue"].round(2)
show("11. PAYMENT METHODS", pay)

# ---------- 12. IQR outliers ----------
q1, q3 = np.percentile(df["Revenue"], [25, 75])
iqr = q3 - q1
lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
out = df[(df["Revenue"] < lower) | (df["Revenue"] > upper)]
show("12. OUTLIER DETECTION (IQR)",
     f"Q1 = {q1:,.2f}\nQ3 = {q3:,.2f}\nIQR = {iqr:,.2f}\n"
     f"Lower bound = {lower:,.2f}\nUpper bound = {upper:,.2f}\nOutliers = {len(out)}")
if len(out):
    show("12b. LARGEST OUTLIER ORDERS",
         out.sort_values("Revenue", ascending=False).head(8)[
             ["Order_ID", "City", "Product", "Category", "Quantity", "Unit_Price", "Revenue"]].round(2))

# ---------- 13. Category summary ----------
cat = (df.groupby("Category")
         .agg(Orders=("Order_ID", "nunique"),
              Revenue=("Revenue", "sum"),
              Qty=("Quantity", "sum"))
         .sort_values("Revenue", ascending=False))
cat["Share_%"] = (cat["Revenue"] / cat["Revenue"].sum() * 100).round(1)
cat["Revenue"] = cat["Revenue"].round(2)
show("13. CATEGORY SUMMARY", cat)

# ---------- 14. NumPy statistics ----------
rev = df["Revenue"].to_numpy()
show("14. NUMPY STATISTICS (Revenue)",
     f"mean   = {np.mean(rev):,.2f}\n"
     f"median = {np.median(rev):,.2f}\n"
     f"std    = {np.std(rev, ddof=1):,.2f}\n"
     f"min    = {np.min(rev):,.2f}\n"
     f"max    = {np.max(rev):,.2f}\n"
     f"p90    = {np.percentile(rev, 90):,.2f}")

# ---------- 15. Weekday pattern ----------
order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
wd = (df.groupby("Weekday")["Revenue"].sum().reindex(order).round(2).to_frame())
show("15. REVENUE BY WEEKDAY", wd)

# ---------- Save cleaned csv ----------
clean_path = r"C:\Users\Lenovo\Documents\Retail-Sales-Analysis\data\retail_sales_clean.csv"
df.drop(columns=["Calc_Revenue"]).to_csv(clean_path, index=False)
show("16. CLEANED FILE SAVED", clean_path)

with open(OUT, "w", encoding="utf8") as f:
    f.write("\n".join(lines))
print(f"\nResults saved -> {OUT}")
