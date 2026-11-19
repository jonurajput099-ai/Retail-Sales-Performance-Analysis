"""Builds and executes analysis.ipynb for the Retail Sales project."""
import nbformat as nbf
from nbclient import NotebookClient

NB = r"C:\Users\Lenovo\Documents\Retail-Sales-Analysis\src\analysis.ipynb"
DATA = r"C:\Users\Lenovo\Documents\Retail-Sales-Analysis\data\retail_sales.csv"

cells = []
def md(text): cells.append(nbf.v4.new_markdown_cell(text.strip("\n")))
def code(text): cells.append(nbf.v4.new_code_cell(text.strip("\n")))

md(r"""
# Retail Sales Performance Analysis
**Project 1 — Python, NumPy & Pandas**

| | |
|---|---|
| **Student** | ami2321 |
| **Dataset** | `data/retail_sales.csv` (502 rows, 10 columns) |
| **Tools** | Python 3.13, NumPy, Pandas |
| **Industry** | Retail / E-commerce |

**Business question:** How is the business performing across products, cities, customers,
payment modes and months — and where should management focus next?

> All numbers below are calculated from the dataset (nothing hard-coded).
""")

md("## 1. Load the dataset and inspect it")
code(r"""
import numpy as np
import pandas as pd

pd.set_option("display.width", 120)
pd.set_option("display.max_columns", 20)

df = pd.read_csv(r"C:\Users\Lenovo\Documents\Retail-Sales-Analysis\data\retail_sales.csv")
print("Rows :", len(df))
print("Cols :", len(df.columns))
df.dtypes.to_frame("dtype")
""")
code("df.head()")
code(r"""
# basic structure
df.info()
""")

md(r"""
## 2. Missing values and duplicate orders

**Observation:** only `City` has missing values (12 rows). Two `Order_ID`s appear twice
(`ORD00010`, `ORD00025`) — exact duplicate records, so they are removed.
""")
code(r"""
missing = df.isna().sum()
print(missing[missing > 0])

dups = df[df.duplicated(subset=["Order_ID"], keep=False)].sort_values("Order_ID")
print("\nDuplicate Order_ID rows:", len(dups))
dups
""")

md(r"""
## 3. Clean the missing `City` values

**Strategy:** missing cities are filled with the **most frequent city inside the same
`Category`** (mode by group), falling back to the global mode. This keeps the city
distribution realistic instead of dropping 12 useful sales rows.
""")
code(r"""
global_mode = df["City"].mode()[0]
cat_mode = df.groupby("Category")["City"].agg(lambda s: s.mode().iat[0] if len(s.mode()) else global_mode)

mask = df["City"].isna()
df.loc[mask, "City"] = df.loc[mask, "Category"].map(cat_mode).fillna(global_mode)
print("filled:", mask.sum(), "rows")

# remove the exact duplicate orders found in step 2
df = df.drop_duplicates(subset=["Order_ID"], keep="first").copy()
print("rows after dedup:", len(df))
df["City"].value_counts()
""")

md("## 4. Convert `Order_Date` and extract month / weekday")
code(r"""
df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")
df["Month"] = df["Order_Date"].dt.month
df["Month_Name"] = df["Order_Date"].dt.month_name()
df["Weekday"] = df["Order_Date"].dt.day_name()

print("Range:", df["Order_Date"].min().date(), "to", df["Order_Date"].max().date())
df[["Order_ID", "Order_Date", "Month", "Month_Name", "Weekday"]].head()
""")

md(r"""
## 5. Row-wise revenue and verification

`Revenue` should equal `Quantity × Unit_Price`. The check below proves whether the supplied
column is trustworthy.
""")
code(r"""
df["Calc_Revenue"] = df["Quantity"] * df["Unit_Price"]
diff = (df["Calc_Revenue"] - df["Revenue"]).abs()

print("max difference      :", round(diff.max(), 4))
print("rows differing >0.01:", int((diff > 0.01).sum()))
print("=> supplied Revenue matches Quantity x Unit_Price" if diff.max() < 0.01 else "=> mismatch found")
""")

md("## 6. Total revenue, total quantity and average order value")
code(r"""
summary = pd.DataFrame({
    "Metric": ["Total revenue", "Total quantity", "Unique orders", "Average order value",
               "Median order value", "Largest single order"],
    "Value": [
        df["Revenue"].sum(),
        df["Quantity"].sum(),
        df["Order_ID"].nunique(),
        df["Revenue"].mean(),
        df["Revenue"].median(),
        df["Revenue"].max(),
    ],
})
summary["Value"] = summary["Value"].map(lambda v: f"{v:,.2f}")
summary
""")

md("## 7. Top 10 products by revenue")
code(r"""
top_products = (df.groupby("Product")
                  .agg(Revenue=("Revenue", "sum"), Units=("Quantity", "sum"), Orders=("Order_ID", "nunique"))
                  .sort_values("Revenue", ascending=False))
top_products["Share_%"] = (top_products["Revenue"] / top_products["Revenue"].sum() * 100).round(1)
top_products.head(10).round(2)
""")

md("## 8. Cities ranked by revenue")
code(r"""
city = (df.groupby("City")
          .agg(Revenue=("Revenue", "sum"), Orders=("Order_ID", "nunique"), Units=("Quantity", "sum"))
          .sort_values("Revenue", ascending=False))
city["Avg_Order_Value"] = (city["Revenue"] / city["Orders"]).round(2)
city["Share_%"] = (city["Revenue"] / city["Revenue"].sum() * 100).round(1)
city.round(2)
""")

md("## 9. Monthly revenue trend and the best month")
code(r"""
monthly = (df.groupby(["Month", "Month_Name"])["Revenue"]
             .sum().reset_index().sort_values("Month"))
monthly["Share_%"] = (monthly["Revenue"] / monthly["Revenue"].sum() * 100).round(1)
monthly["Revenue"] = monthly["Revenue"].round(2)

best = monthly.loc[monthly["Revenue"].idxmax()]
print(f"Best month: {best['Month_Name']} -> {best['Revenue']:,.2f}")
monthly
""")

md("## 10. Top 10 customers by total spending")
code(r"""
customers = (df.groupby("Customer_ID")
               .agg(Revenue=("Revenue", "sum"), Orders=("Order_ID", "nunique"))
               .sort_values("Revenue", ascending=False))
customers["Avg_Order_Value"] = (customers["Revenue"] / customers["Orders"]).round(2)
customers.head(10).round(2)
""")

md("## 11. Payment methods: orders vs revenue")
code(r"""
pay = (df.groupby("Payment_Mode")
         .agg(Orders=("Order_ID", "nunique"), Revenue=("Revenue", "sum"))
         .sort_values("Revenue", ascending=False))
pay["Share_%"] = (pay["Revenue"] / pay["Revenue"].sum() * 100).round(1)
pay["Avg_Order_Value"] = (pay["Revenue"] / pay["Orders"]).round(2)
pay.round(2)
""")

md(r"""
## 12. Unusually large orders (IQR method)

`upper = Q3 + 1.5 × IQR`. Orders above this line are statistically unusual and should be
checked by the sales team (possible bulk / corporate orders).
""")
code(r"""
q1, q3 = np.percentile(df["Revenue"], [25, 75])
iqr = q3 - q1
lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr

outliers = df[(df["Revenue"] < lower) | (df["Revenue"] > upper)]
print(f"Q1={q1:,.2f}  Q3={q3:,.2f}  IQR={iqr:,.2f}")
print(f"Bounds: {lower:,.2f} .. {upper:,.2f}")
print("Outlier orders:", len(outliers), f"({len(outliers)/len(df)*100:.1f}% of orders)")
print("Outlier revenue share: {:.1f}%".format(outliers["Revenue"].sum() / df["Revenue"].sum() * 100))
outliers.nlargest(8, "Revenue")[["Order_ID", "City", "Product", "Quantity", "Unit_Price", "Revenue"]]
""")

md("## 13. Product-category summary (`groupby`)")
code(r"""
cat = (df.groupby("Category")
         .agg(Orders=("Order_ID", "nunique"), Revenue=("Revenue", "sum"), Units=("Quantity", "sum"))
         .sort_values("Revenue", ascending=False))
cat["Share_%"] = (cat["Revenue"] / cat["Revenue"].sum() * 100).round(1)
cat["Avg_Order_Value"] = (cat["Revenue"] / cat["Orders"]).round(2)
cat.round(2)
""")

md(r"""
### 13.1 Data-quality check: is `Category` a product attribute?

Before drawing conclusions from the category table, the analyst must confirm that each product
belongs to exactly one category. If a product appears under both categories, `Category` is only a
**transaction-level label**, and category results describe *where revenue was booked*, not the
product line itself.
""")
code(r"""
check = pd.DataFrame({
    "distinct_categories": df.groupby("Product")["Category"].nunique(),
    "categories_seen": df.groupby("Product")["Category"].agg(lambda s: ", ".join(sorted(s.unique()))),
    "distinct_prices": df.groupby("Product")["Unit_Price"].nunique(),
    "min_price": df.groupby("Product")["Unit_Price"].min(),
    "max_price": df.groupby("Product")["Unit_Price"].max(),
})
check
""")
code(r"""
inconsistent = df.groupby("Product")["Category"].nunique().gt(1).sum()
print("products mapped to more than one category:", inconsistent, "of", df["Product"].nunique())
if inconsistent:
    print("=> Category is a transaction label, not a fixed product attribute.")
    print("   Category shares are valid for revenue reporting, but must NOT be read as")
    print("   'this product line belongs to this category'.")
print("\nprice spread per product: every product sells between",
      f"{df['Unit_Price'].min():,} and {df['Unit_Price'].max():,}",
      "=> price positioning per product cannot be inferred from this synthetic dataset.")
""")

md("## 14. NumPy statistics on order revenue")
code(r"""
rev = df["Revenue"].to_numpy()

stats = {
    "mean": np.mean(rev),
    "median": np.median(rev),
    "std (sample)": np.std(rev, ddof=1),
    "min": np.min(rev),
    "max": np.max(rev),
    "25th pct": np.percentile(rev, 25),
    "75th pct": np.percentile(rev, 75),
    "90th pct": np.percentile(rev, 90),
}
pd.Series(stats).round(2).to_frame("value")
""")

md("## 15. Weekday pattern")
code(r"""
order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
wd = (df.groupby("Weekday")
        .agg(Revenue=("Revenue", "sum"), Orders=("Order_ID", "nunique"))
        .reindex(order))
wd["Avg_Order_Value"] = (wd["Revenue"] / wd["Orders"]).round(2)
wd.round(2)
""")

md(r"""
## 16. Save the cleaned dataset

The cleaned file is written to `data/retail_sales_clean.csv` so every number above can be
reproduced.
""")
code(r"""
clean_path = r"C:\Users\Lenovo\Documents\Retail-Sales-Analysis\data\retail_sales_clean.csv"
df.drop(columns=["Calc_Revenue"]).to_csv(clean_path, index=False)
print("saved:", clean_path, "| rows:", len(df))
""")

md(r"""
## 17. Five business recommendations (evidence-based)

1. **Protect the top revenue products.** `Keyboard`, `Headphones` and `Laptop` are the highest
   revenue generators — keep stock availability high and bundle accessories with them.
2. **Focus marketing on Delhi and Bengaluru.** They contribute the largest revenue share, and
   Delhi also has the highest average order value, so premium offers fit there best.
3. **Plan for the February–April peak and June–July dip.** Reallocate ad spend to strong months
   and run clearance/promo campaigns in the low months (June revenue share is only ~3%).
4. **Grow UPI and Card share.** UPI already leads with ~30% revenue share; cash-on-delivery still
   holds ~26% — cashback offers on UPI/Card reduce handling cost and improve prepaid ratio.
5. **Audit the 32 statistically unusual orders.** They represent a disproportionate revenue share;
   confirming they are genuine bulk/corporate orders (not data errors) protects reporting quality,
   and those customers deserve a key-account treatment.

### Limitations
- The dataset is synthetic and covers ~17 months; seasonality conclusions should be re-validated
  with more history.
- Missing `City` values were imputed (12 rows), which slightly smooths city-level results.
- `Category` is **not** a fixed product attribute in this dataset: all six products appear under
  both Accessories and Electronics, and every product sells across the full price range
  (₹499–₹45,999). Category shares are therefore valid for revenue reporting, but must not be read
  as product-line membership or price positioning.
- Correlation between variables is not causal; recommendations are directional, not guaranteed.
""")

nb = nbf.v4.new_notebook()
nb["cells"] = cells
nb["metadata"] = {
    "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
    "language_info": {"name": "python", "version": "3.13"},
}

print("executing notebook...")
client = NotebookClient(nb, timeout=600, kernel_name="python3", resources={"metadata": {"path": r"C:\Users\Lenovo\Documents\Retail-Sales-Analysis\src"}})
client.execute()

with open(NB, "w", encoding="utf8") as f:
    nbf.write(nb, f)
print("saved:", NB)
