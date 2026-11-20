# Retail Sales Performance Analysis

**Project 1 — Python, NumPy & Pandas (Individual Industry Project)**
Industry: Retail / E-commerce · Dataset: `01_retail_sales.csv` (502 records × 10 columns)

**Prepared by:** Khemchand Singh

An end-to-end, reproducible analysis of a retail transaction export: cleaning the raw data,
validating revenue, and turning it into evidence-based commercial insight across products,
cities, customers, payment modes and time.

---

## Business question

> How is the business performing across products, cities, customers, payment modes and months —
> and where should management focus next?

Every number in the notebook is **computed from the dataset**, never typed in manually.

---

## Key findings

| Metric | Result |
|---|---|
| Total revenue | **₹ 1,39,54,120** |
| Total quantity sold | 1,480 units |
| Unique orders (after cleaning) | 500 |
| Average order value | ₹ 27,908.24 |
| Median order value | ₹ 4,998 |
| Largest single order | ₹ 2,29,995 |
| Top product | **Keyboard** — ₹ 34,38,656 (24.6% of revenue) |
| Top city | **Delhi** — ₹ 33,97,468 (24.3%), highest AOV ₹ 30,335 |
| Best month | **February** — ₹ 20,03,009 (14.4%) |
| Weakest month | **June** — ₹ 4,31,919 (3.1%) |
| Top customer | C0018 — ₹ 4,45,986 across 4 orders |
| Leading payment mode | **UPI** — ₹ 42,46,098 (30.4% of revenue) |
| Outliers (1.5 × IQR) | 32 orders above ₹ 93,897 → **41.5% of all revenue** |
| Largest category | **Accessories** — ₹ 74,92,779 (53.7%) |
| Revenue distribution | mean ₹ 27,908 · median ₹ 4,998 · std ₹ 47,747 |

**Five recommendations (short version)**

1. **Protect the top products** — Keyboard, Headphones and Laptop drive 61.8% of revenue; keep them in stock and bundle accessories.
2. **Focus marketing on Delhi and Bengaluru** — largest revenue share, and Delhi has the highest average order value (₹ 30,335).
3. **Plan around the Feb–Apr peak and Jun–Jul dip** — June contributes only 3.1%, so run clearance campaigns there.
4. **Push UPI and Card** — UPI already leads at 30.4%; cash is still 25.8% — cashback on prepaid channels lowers handling cost.
5. **Audit the 32 high-value orders** — they carry 41.5% of revenue; confirm they are genuine bulk/corporate orders and give those customers key-account treatment.

---

## Data quality work

| Issue | Extent | Treatment |
|---|---|---|
| Missing `City` | 12 rows (2.4%) | Imputed with the **most frequent city within the same `Category`**, falling back to the global mode — keeps the 12 revenue-bearing rows instead of dropping them. |
| Duplicate orders | 2 `Order_ID`s (4 rows) | Exact duplicates removed with `drop_duplicates(subset="Order_ID", keep="first")` → **502 → 500 rows**. |
| `Order_Date` as text | All rows | Converted with `pd.to_datetime(errors="coerce")`; `Month`, `Month_Name`, `Weekday` extracted. |
| `Revenue` trustworthiness | All rows | Recalculated as `Quantity × Unit_Price`; **maximum difference = 0** → supplied column validated. |
| Extreme order values | 32 orders | Flagged with the 1.5 × IQR rule (upper bound ₹ 93,897) and **retained** — plausible bulk orders, not data errors. Their revenue share is reported explicitly. |
| `Category` not tied to `Product` | All 6 products | Every product appears under **both** Accessories and Electronics, and each product sells across the full ₹499–₹45,999 range. `Category` is therefore treated as a **transaction-level label**: its revenue shares are reported, but no product-line or price-positioning claim is made (verified in notebook §13.1). |

---

## Repository structure

```
Retail-Sales-Analysis/
├── README.md
├── requirements.txt
├── SRS/
│   └── SRS_Retail_Sales_Performance_Analysis.docx   # 12-section SRS (PDF/DOCX submission)
├── src/
│   ├── analysis.ipynb          # executed notebook — all outputs visible
│   ├── analysis.py             # plain-Python equivalent (writes results.txt)
│   ├── results.txt             # console output of analysis.py
│   ├── build_notebook.py       # regenerates + executes analysis.ipynb
│   └── build_srs.py            # regenerates the SRS .docx
└── data/
    ├── retail_sales.csv        # raw dataset (never modified)
    └── retail_sales_clean.csv  # cleaned output (500 rows)
```

---

## How to run

### 1. Install dependencies

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

### 2. Run the analysis

**Option A — notebook (recommended, shows all tables)**

```bash
jupyter notebook src/analysis.ipynb
```
Run all cells top-to-bottom. The notebook writes `data/retail_sales_clean.csv`.

**Option B — plain script**

```bash
python src/analysis.py
```
Prints every table to the console and writes `src/results.txt`.

### 3. Regenerate the deliverables (optional)

```bash
python src/build_notebook.py   # rebuilds and executes analysis.ipynb
python src/build_srs.py        # rebuilds the SRS document
```

---

## Method

1. **Acquire** — load CSV, inspect shape, dtypes, `head()`, `info()`.
2. **Assess quality** — missing values per column, duplicated `Order_ID`s, date parsing.
3. **Clean** — impute `City`, drop duplicates, convert dates, engineer `Month`/`Weekday`, validate revenue.
4. **Describe** — `groupby` aggregations: totals, top-10 products, city ranking, monthly trend, top-10 customers, payment comparison, category summary, weekday pattern.
5. **Analyse statistically** — NumPy mean/median/std/percentiles; 1.5 × IQR outlier detection.
6. **Interpret** — convert findings into five actionable recommendations; state assumptions and limitations.

---

## Requirements coverage (brief → code)

| # | Brief task | Notebook section |
|---|---|---|
| 1 | Load and inspect dimensions, columns, dtypes | §1 |
| 2 | Missing values and duplicate orders | §2 |
| 3 | Clean missing `City` with justified strategy | §3 |
| 4 | Convert `Order_Date`, extract month & weekday | §4 |
| 5 | Row-wise revenue + verify supplied `Revenue` | §5 |
| 6 | Total revenue, quantity, average order value | §6 |
| 7 | Top 10 products by revenue | §7 |
| 8 | Cities ranked by revenue | §8 |
| 9 | Monthly revenue + best month | §9 |
| 10 | Top 10 customers by spending | §10 |
| 11 | Payment methods: orders vs revenue | §11 |
| 12 | IQR outlier detection | §12 |
| 13 | Product-category summary (`groupby`) | §13 |
| 14 | NumPy mean, median, standard deviation | §14 |
| 15 | Five business recommendations | §17 |
| + | Weekday pattern (extra insight) | §15 |

---

## Assumptions & limitations

- The dataset is **synthetic** — absolute revenue does not represent a real company.
- `Category` is not a fixed product attribute: every product appears under both Accessories and Electronics, and every product spans the full price range. Category revenue shares are reported as transaction labels only.
- Imputing 12 city values slightly smooths city-level results (documented, not hidden).
- ~17 months of history: seasonality conclusions should be re-validated with more data.
- Analysis is **descriptive**; no causal claim is made between payment mode, city, product and revenue.

---

## Tech stack

| Tool | Version |
|---|---|
| Python | 3.13 |
| pandas | 3.1.1 |
| numpy | 2.5.5 |
| nbformat / nbclient | 5.11 / 0.11 |
| python-docx | 1.2.0 |

---

## Author

**Khemchand Singh** — Five-Day Engineering Student Training, First Year / First Semester.

Submitted with: SRS document (DOCX + PDF), executed notebook, and this GitHub repository.
