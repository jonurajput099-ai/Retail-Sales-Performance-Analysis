"""
Generates the SRS (Software Requirements Specification) document for
Project 1 - Retail Sales Performance Analysis.

Output: SRS/SRS_Retail_Sales_Performance_Analysis.docx
"""
import os
from datetime import date

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = r"C:\Users\Lenovo\Documents\Retail-Sales-Analysis"
OUT = os.path.join(ROOT, "SRS", "SRS_Retail_Sales_Performance_Analysis.docx")

ACCENT = RGBColor(0x1F, 0x4E, 0x79)
GREY = RGBColor(0x59, 0x59, 0x59)


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def shade(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False, size=9, color=None, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    if color:
        run.font.color.rgb = color
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def add_table(doc, headers, rows, widths=None, header_fill="1F4E79",
              font_size=9, zebra=True, first_col_bold=False):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_text(hdr[i], h, bold=True, size=font_size,
                      color=RGBColor(0xFF, 0xFF, 0xFF))
        shade(hdr[i], header_fill)

    for r_idx, row in enumerate(rows):
        cells = table.add_row().cells
        for c_idx, val in enumerate(row):
            set_cell_text(cells[c_idx], val, size=font_size,
                          bold=(first_col_bold and c_idx == 0))
            if zebra and r_idx % 2 == 1:
                shade(cells[c_idx], "F2F5F9")

    if widths:
        for row in table.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)

    # repeat header row on page breaks
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return table


def bullet(doc, text, level=0):
    p = doc.add_paragraph(text, style="List Bullet" if level == 0 else "List Bullet 2")
    p.paragraph_format.space_after = Pt(2)
    return p


def numbered(doc, text):
    p = doc.add_paragraph(text, style="List Number")
    p.paragraph_format.space_after = Pt(2)
    return p


def para(doc, text, size=10.5, bold=False, italic=False, space_after=6,
         align=WD_ALIGN_PARAGRAPH.JUSTIFY, color=None):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    if color:
        run.font.color.rgb = color
    return p


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = ACCENT
        run.font.name = "Calibri"
        run.font.size = Pt(15 if level == 1 else 12.5 if level == 2 else 11)
    h.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    h.paragraph_format.space_after = Pt(4)
    return h


def add_page_number_footer(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run("Khemchand Singh  |  Retail Sales Performance Analysis  |  Page ")
    run.font.size = Pt(8)
    run.font.color.rgb = GREY

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run2 = p.add_run()
    run2.font.size = Pt(8)
    run2.font.color.rgb = GREY
    run2._r.append(fld_begin)
    run2._r.append(instr)
    run2._r.append(fld_end)

    run3 = p.add_run(" of ")
    run3.font.size = Pt(8)
    run3.font.color.rgb = GREY

    fld_begin2 = OxmlElement("w:fldChar")
    fld_begin2.set(qn("w:fldCharType"), "begin")
    instr2 = OxmlElement("w:instrText")
    instr2.set(qn("xml:space"), "preserve")
    instr2.text = "NUMPAGES"
    fld_end2 = OxmlElement("w:fldChar")
    fld_end2.set(qn("w:fldCharType"), "end")

    run4 = p.add_run()
    run4.font.size = Pt(8)
    run4.font.color.rgb = GREY
    run4._r.append(fld_begin2)
    run4._r.append(instr2)
    run4._r.append(fld_end2)


def hrule(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "1F4E79")
    pbdr.append(bottom)
    pPr.append(pbdr)


# --------------------------------------------------------------------------- #
# document
# --------------------------------------------------------------------------- #
doc = Document()

# base style
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")

section = doc.sections[0]
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(1.8)
section.left_margin = Cm(2.2)
section.right_margin = Cm(2.0)
add_page_number_footer(section)

# ---------------------------- COVER PAGE ---------------------------- #
for _ in range(2):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("SOFTWARE REQUIREMENTS SPECIFICATION (SRS)")
r.bold = True
r.font.size = Pt(11)
r.font.color.rgb = GREY
p.paragraph_format.space_after = Pt(2)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("RETAIL SALES PERFORMANCE ANALYSIS")
r.bold = True
r.font.size = Pt(26)
r.font.color.rgb = ACCENT
p.paragraph_format.space_after = Pt(2)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Python, NumPy & Pandas — Individual Industry Project (Project 1)")
r.font.size = Pt(12)
r.italic = True
r.font.color.rgb = GREY
p.paragraph_format.space_after = Pt(18)

hrule(doc)

add_table(
    doc,
    ["Field", "Details"],
    [
        ["Project title", "Retail Sales Performance Analysis"],
        ["Project number", "Project 1 of 30"],
        ["Industry / domain", "Retail / E-commerce"],
        ["Dataset", "01_retail_sales.csv (502 records, 10 columns)"],
        ["Primary tools", "Python 3.13, NumPy, Pandas, Jupyter Notebook"],
        ["Document version", "1.0"],
        ["Prepared by", "Khemchand Singh"],
        ["Course", "Five-Day Engineering Student Training — First Year, First Semester"],
        ["Date of submission", date.today().strftime("%d %B %Y")],
        ["Repository", "github.com/jonurajput099-ai/Retail-Sales-Performance-Analysis"],
    ],
    widths=[5.0, 11.0],
    header_fill="1F4E79",
    font_size=10,
    first_col_bold=True,
)

para(doc, "", space_after=10)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(10)
r = p.add_run("Prepared by:  KHEMCHAND SINGH")
r.bold = True
r.font.size = Pt(13)
r.font.color.rgb = ACCENT
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("This document defines the problem, objectives, scope, dataset, functional "
              "requirements, methodology and expected outcomes of the project. It is submitted "
              "together with the source code (Jupyter Notebook) and the GitHub repository link.")
r.font.size = Pt(9.5)
r.italic = True
r.font.color.rgb = GREY

doc.add_page_break()

# ---------------------------- CONTENTS ---------------------------- #
heading(doc, "Table of Contents", 1)
toc = [
    "1. Introduction",
    "    1.1 Background",
    "    1.2 Problem Statement",
    "    1.3 Purpose of this Document",
    "2. Objectives",
    "3. Scope",
    "4. Dataset Description",
    "5. Functional Requirements",
    "6. Non-Functional Requirements",
    "7. Tools and Technologies",
    "8. Methodology",
    "9. Expected Outcomes and Deliverables",
    "10. Assumptions, Constraints and Limitations",
    "11. Project Plan",
    "12. Conclusion",
    "Appendix A — Result Snapshot",
    "Appendix B — References",
]
for item in toc:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.6)
    r = p.add_run(item)
    r.font.size = Pt(10.5)
    if not item.startswith("    "):
        r.bold = True

doc.add_page_break()

# ---------------------------- 1. INTRODUCTION ---------------------------- #
heading(doc, "1. Introduction", 1)

heading(doc, "1.1 Background", 2)
para(doc,
     "Retail and e-commerce businesses generate a large volume of transaction records every day. "
     "Each record captures what was sold, where, to whom, in what quantity, at what price and "
     "through which payment channel. Individually these rows are meaningless, but when cleaned and "
     "aggregated they describe the commercial health of the business: which products earn money, "
     "which cities perform, which months are strong or weak, which customers are most valuable and "
     "which payment methods dominate.")
para(doc,
     "Historically this work was done with spreadsheets. Modern entry-level analyst roles expect the "
     "same work to be performed programmatically, because code is reproducible, auditable and scales "
     "to millions of rows. This project applies the Python data stack — Pandas for tabular "
     "manipulation and NumPy for numerical statistics — to a retail transaction dataset containing "
     "502 order records.")

heading(doc, "1.2 Problem Statement", 2)
para(doc,
     "The retail business in this scenario maintains a raw order-level CSV export. The export is "
     "not analysis-ready: it contains missing city values and duplicated order records, it stores "
     "dates as plain text, and it offers no summarised view of performance. Management currently "
     "cannot answer basic commercial questions without manual spreadsheet work:")
bullet(doc, "What is the total revenue, total quantity sold and average order value?")
bullet(doc, "Which products, cities, customers and payment methods generate the most revenue?")
bullet(doc, "How does revenue behave across months, and which month performs best?")
bullet(doc, "Which orders are statistical outliers and might distort reporting?")
bullet(doc, "Which category (Accessories or Electronics) is more valuable?")
para(doc,
     "In addition, the supplied Revenue column must be validated against Quantity × Unit_Price "
     "before any conclusion is drawn from it. Without a documented, reproducible analysis pipeline, "
     "the business risks making stocking, marketing and payment-incentive decisions on unreliable "
     "numbers.")

heading(doc, "1.3 Purpose of this Document", 2)
para(doc,
     "This SRS specifies the requirements of the Retail Sales Performance Analysis project. It "
     "defines the business problem, the analytical objectives, the scope of work, the dataset and "
     "its quality issues, the functional and non-functional requirements, the tools used, the "
     "methodology followed, and the deliverables produced. It is intended to be read together with "
     "the executed Jupyter Notebook (analysis.ipynb) that implements every requirement listed in "
     "Section 5.")

# ---------------------------- 2. OBJECTIVES ---------------------------- #
heading(doc, "2. Objectives", 1)

heading(doc, "2.1 General Objective", 2)
para(doc,
     "To build a reproducible Python-based analysis pipeline that cleans the raw retail transaction "
     "export and converts it into evidence-based commercial insight covering revenue, product, "
     "city, customer, payment and time-based performance.")

heading(doc, "2.2 Specific Objectives", 2)
objectives = [
    ("O1", "Inspect the dataset structure — dimensions, column names and data types — before any transformation."),
    ("O2", "Detect and treat data-quality problems: missing values, duplicated orders and incorrect data types."),
    ("O3", "Validate the integrity of the Revenue column against Quantity × Unit_Price."),
    ("O4", "Produce core business KPIs: total revenue, total quantity, unique orders and average order value."),
    ("O5", "Rank products, cities and customers to identify the highest-value contributors."),
    ("O6", "Analyse revenue over time (monthly trend, best month, weekday pattern)."),
    ("O7", "Compare payment methods by both order count and revenue contribution."),
    ("O8", "Apply the IQR method to detect statistically unusual orders and quantify their revenue impact."),
    ("O9", "Use NumPy to compute mean, median, standard deviation and percentiles of order revenue."),
    ("O10", "Translate every finding into five concrete, evidence-backed business recommendations."),
]
add_table(doc, ["ID", "Objective"], objectives, widths=[1.6, 14.4], font_size=9.5)

# ---------------------------- 3. SCOPE ---------------------------- #
heading(doc, "3. Scope", 1)

heading(doc, "3.1 In Scope", 2)
bullet(doc, "Loading the supplied 01_retail_sales.csv file with Pandas and profiling its structure.")
bullet(doc, "Cleaning: imputing 12 missing City values with a documented per-category mode strategy, "
            "removing exact duplicate orders, converting Order_Date to a proper datetime type.")
bullet(doc, "Feature engineering: Month, Month_Name and Weekday columns; a calculated revenue column "
            "for validation.")
bullet(doc, "Descriptive analytics on revenue, quantity, orders, products, cities, customers, "
            "payment modes, categories, months and weekdays.")
bullet(doc, "Outlier detection on order revenue using the 1.5 × IQR rule and NumPy percentiles.")
bullet(doc, "Export of a cleaned dataset (retail_sales_clean.csv) so every reported number can be reproduced.")
bullet(doc, "Five business recommendations with the supporting numbers behind each one.")

heading(doc, "3.2 Out of Scope", 2)
bullet(doc, "Predictive modelling, forecasting and machine learning.")
bullet(doc, "Causal inference — observed relationships are described, not proven causal.")
bullet(doc, "Live database connectivity; the analysis runs on a static CSV export.")
bullet(doc, "Interactive dashboards (Power BI / Tableau) and production deployment.")
bullet(doc, "Financial accounting treatment (tax, returns, refunds, cost of goods sold).")

# ---------------------------- 4. DATASET ---------------------------- #
heading(doc, "4. Dataset Description", 1)
para(doc,
     "The dataset is a synthetic teaching dataset supplied with the project brief. It represents "
     "order-level retail transactions. File: 01_retail_sales.csv, 502 rows × 10 columns, "
     "approximately 35 KB.")

heading(doc, "4.1 Column Dictionary", 2)
add_table(
    doc,
    ["Column", "Data type", "Description", "Analytical use"],
    [
        ["Order_ID", "Text", "Unique order identifier (ORDxxxxx)", "Primary key; duplicate detection"],
        ["Order_Date", "Text → Date", "Date the order was placed", "Month/weekday trends"],
        ["City", "Text", "Customer city (12 missing)", "Regional performance ranking"],
        ["Product", "Text", "Product purchased", "Top-product revenue ranking"],
        ["Category", "Text", "Accessories or Electronics", "Category grouping"],
        ["Quantity", "Integer", "Units purchased per order", "Total quantity; revenue validation"],
        ["Unit_Price", "Integer", "Price per unit (₹)", "Revenue validation"],
        ["Customer_ID", "Text", "Customer identifier (Cxxxx)", "Customer value ranking"],
        ["Payment_Mode", "Text", "UPI, Card, Cash, NetBanking", "Payment channel comparison"],
        ["Revenue", "Integer", "Order revenue (₹)", "Primary KPI; outlier detection"],
    ],
    widths=[2.6, 2.2, 5.6, 5.6],
    font_size=8.5,
    first_col_bold=True,
)

heading(doc, "4.2 Data Quality Assessment", 2)
add_table(
    doc,
    ["Issue", "Extent", "Treatment applied"],
    [
        ["Missing City values", "12 of 502 rows (2.4%)",
         "Imputed with the most frequent city within the same Category, falling back to the global mode. "
         "Preserves 12 revenue-bearing rows instead of discarding them."],
        ["Duplicated Order_ID records", "2 Order_IDs repeated (4 rows total)",
         "Exact duplicate records; second copy removed with drop_duplicates(subset='Order_ID', keep='first'). "
         "Final row count 500."],
        ["Order_Date stored as text", "All 502 rows",
         "Converted with pd.to_datetime(errors='coerce'); Month, Month_Name and Weekday extracted."],
        ["Revenue column not verified", "All rows",
         "Recalculated as Quantity × Unit_Price and compared; maximum absolute difference = 0, "
         "confirming the supplied column is correct."],
        ["Potential extreme values", "32 orders above the IQR upper bound",
         "Flagged and quantified (41.5% of total revenue); retained because they are plausible "
         "bulk/high-value orders, not data-entry errors."],
        ["Category not tied to Product", "All 6 products appear in both categories",
         "Detected and documented. Each product (Headphones, Keyboard, Laptop, Monitor, Mouse, Webcam) "
         "is recorded under both Accessories and Electronics, and every product sells across the full "
         "₹499–₹45,999 price range. Category is therefore treated as a transaction-level label: its "
         "revenue shares are reported, but no product-line or price-positioning claim is made."],
    ],
    widths=[3.6, 3.6, 8.8],
    font_size=8.5,
    first_col_bold=True,
)

# ---------------------------- 5. FUNCTIONAL REQUIREMENTS ---------------------------- #
heading(doc, "5. Functional Requirements", 1)
para(doc,
     "Each requirement below maps directly to a task in the project brief. Priority: H = High "
     "(must work for submission), M = Medium (expected), L = Low (enhancement). Every requirement "
     "is implemented as a labelled section of analysis.ipynb and produces visible output.")

frs = [
    ["FR-01", "The system shall load 01_retail_sales.csv into a Pandas DataFrame and report the "
              "number of rows, number of columns, column names and data types.", "H"],
    ["FR-02", "The system shall report the count of missing values per column and identify "
              "duplicated Order_ID records.", "H"],
    ["FR-03", "The system shall impute missing City values using a justified strategy "
              "(per-category mode) and report how many rows were filled.", "H"],
    ["FR-04", "The system shall remove exact duplicate orders by Order_ID and report the resulting "
              "row count (502 → 500).", "H"],
    ["FR-05", "The system shall convert Order_Date to datetime and create Month, Month_Name and "
              "Weekday columns.", "H"],
    ["FR-06", "The system shall calculate row-wise revenue as Quantity × Unit_Price and verify it "
              "against the supplied Revenue column, reporting the maximum difference.", "H"],
    ["FR-07", "The system shall compute total revenue, total quantity, unique order count, average "
              "order value, median order value and largest order value.", "H"],
    ["FR-08", "The system shall rank the top 10 products by total revenue, including units sold, "
              "order count and revenue share.", "H"],
    ["FR-09", "The system shall rank all cities by total revenue with order count, units and "
              "average order value.", "H"],
    ["FR-10", "The system shall compute monthly revenue, each month's percentage share, and "
              "identify the best-performing month.", "H"],
    ["FR-11", "The system shall identify the top 10 customers by total spending, with order count "
              "and average order value per customer.", "M"],
    ["FR-12", "The system shall compare payment methods by number of orders, total revenue, revenue "
              "share and average order value.", "H"],
    ["FR-13", "The system shall detect unusually large order values using the 1.5 × IQR rule with "
              "np.percentile, reporting the bounds, count and revenue share of outliers.", "H"],
    ["FR-14", "The system shall produce a product-category summary using groupby with orders, "
              "revenue, units and share.", "H"],
    ["FR-15", "The system shall compute mean, median, standard deviation, min, max and selected "
              "percentiles of order revenue using NumPy.", "H"],
    ["FR-16", "The system shall present five business recommendations, each supported by a specific "
              "result from the analysis.", "H"],
    ["FR-17", "The system shall save the cleaned dataset to retail_sales_clean.csv for "
              "reproducibility.", "M"],
    ["FR-18", "The analysis shall include a weekday revenue pattern comparison as an additional "
              "insight beyond the required tasks.", "L"],
]
add_table(doc, ["ID", "Requirement", "Priority"], frs, widths=[1.6, 13.2, 1.8], font_size=8.5,
          first_col_bold=True)

# ---------------------------- 6. NFR ---------------------------- #
heading(doc, "6. Non-Functional Requirements", 1)
nfrs = [
    ["NFR-01 Reproducibility", "Running the notebook top-to-bottom must regenerate every number and "
                               "the cleaned CSV without manual edits.", "Notebook is executed end-to-end; "
                               "outputs are stored in the file."],
    ["NFR-02 No hard-coded results", "All reported values must be computed from the dataset, never "
                                     "typed manually.", "Every figure comes from a Pandas/NumPy expression."],
    ["NFR-03 Readability", "Code must use meaningful variable names, consistent section headings and "
                           "Markdown explanations of key decisions.", "17 Markdown sections + 18 code cells."],
    ["NFR-04 Performance", "Full pipeline must run in a few seconds on a standard laptop.", "Executes in under 5 seconds."],
    ["NFR-05 Portability", "Must run on Windows, macOS or Linux with Python 3.10+ and the pinned "
                          "packages in requirements.txt.", "Only pandas, numpy, nbformat, nbclient, ipykernel."],
    ["NFR-06 Data safety", "The raw source CSV must never be modified in place.", "Cleaning writes to a "
                                                                             "separate cleaned CSV."],
    ["NFR-07 Documentation", "A README must explain setup, execution and findings for a new reader.",
     "README.md included in the repository."],
]
add_table(doc, ["Requirement", "Description", "How it is met"], nfrs,
          widths=[3.6, 7.0, 5.4], font_size=8.5, first_col_bold=True)

# ---------------------------- 7. TOOLS ---------------------------- #
heading(doc, "7. Tools and Technologies", 1)
add_table(
    doc,
    ["Tool / Library", "Version", "Role in the project"],
    [
        ["Python", "3.13.17", "Programming language for the entire pipeline"],
        ["Pandas", "3.1.1", "Loading CSV, cleaning, groupby aggregation, sorting, datetime handling, export"],
        ["NumPy", "2.5.5", "Percentiles (IQR bounds) and descriptive statistics on revenue arrays"],
        ["Jupyter Notebook / nbformat / nbclient", "5.11 / 0.11", "Authoring and programmatic execution of analysis.ipynb"],
        ["python-docx", "1.2.0", "Generation of this SRS document"],
        ["Git & GitHub", "—", "Version control and public repository hosting for submission"],
    ],
    widths=[5.2, 2.6, 8.2],
    font_size=9,
    first_col_bold=True,
)

# ---------------------------- 8. METHODOLOGY ---------------------------- #
heading(doc, "8. Methodology", 1)
para(doc,
     "The project follows a standard six-phase data-analysis workflow. Each phase produces a visible "
     "artefact, so the reasoning can be followed and audited.")

phases = [
    ["1. Data acquisition",
     "Load the CSV with pd.read_csv, inspect shape, dtypes, head() and info().",
     "Structural profile of 502 rows × 10 columns."],
    ["2. Data quality assessment",
     "Count missing values per column; detect duplicated Order_IDs; test date parsing.",
     "12 missing City values, 2 duplicated orders, dates stored as text."],
    ["3. Data cleaning",
     "Impute City by per-category mode; drop duplicate orders; convert Order_Date to datetime; "
     "create Month, Month_Name, Weekday; recalculate revenue and compare with the supplied column.",
     "Clean dataset of 500 rows; revenue validated (difference = 0)."],
    ["4. Descriptive analysis",
     "Group and aggregate with Pandas: totals, top-10 products, city ranking, monthly trend, top-10 "
     "customers, payment comparison, category summary, weekday pattern.",
     "All KPI and ranking tables."],
    ["5. Statistical analysis",
     "NumPy mean, median, standard deviation, percentiles; 1.5 × IQR outlier rule for order revenue.",
     "Distribution profile and 32 flagged outlier orders."],
    ["6. Interpretation",
     "Convert each finding into a business action; state assumptions and limitations.",
     "Five evidence-based recommendations."],
]
add_table(doc, ["Phase", "Activities", "Output"], phases, widths=[3.2, 8.4, 4.4], font_size=8.5,
          first_col_bold=True)

# ---------------------------- 9. OUTCOMES ---------------------------- #
heading(doc, "9. Expected Outcomes and Deliverables", 1)

heading(doc, "9.1 Analytical Outputs", 2)
bullet(doc, "Revenue summary table — total revenue, total quantity, unique orders, average/median/largest order value.")
bullet(doc, "Top-10 products by revenue, with units, order count and revenue share.")
bullet(doc, "City ranking by revenue with average order value per city.")
bullet(doc, "Monthly revenue trend with percentage share and the identified best month.")
bullet(doc, "Top-10 customers by total spending.")
bullet(doc, "Payment-method comparison by orders, revenue and average order value.")
bullet(doc, "IQR outlier report with bounds, count and revenue share.")
bullet(doc, "Product-category summary via groupby.")
bullet(doc, "NumPy statistical profile of order revenue.")
bullet(doc, "Weekday revenue pattern (additional insight).")
bullet(doc, "Cleaned dataset export (retail_sales_clean.csv).")
bullet(doc, "Five evidence-based business recommendations.")

heading(doc, "9.2 Submission Deliverables", 2)
add_table(
    doc,
    ["Deliverable", "File / location", "Purpose"],
    [
        ["SRS document", "SRS/SRS_Retail_Sales_Performance_Analysis.docx", "This document"],
        ["Source code (notebook)", "src/analysis.ipynb", "Executed analysis with all outputs visible"],
        ["Reproducible script", "src/analysis.py", "Plain-Python equivalent producing results.txt"],
        ["Cleaned dataset", "data/retail_sales_clean.csv", "Auditable cleaned data"],
        ["README", "README.md", "Setup, structure, findings, how to run"],
        ["Requirements", "requirements.txt", "Pinned dependency list"],
        ["GitHub repository", "Public repo link", "Hosts SRS, code, dataset and README"],
    ],
    widths=[4.2, 7.4, 4.4],
    font_size=8.5,
    first_col_bold=True,
)

# ---------------------------- 10. ASSUMPTIONS ---------------------------- #
heading(doc, "10. Assumptions, Constraints and Limitations", 1)

heading(doc, "10.1 Assumptions", 2)
bullet(doc, "Each row represents one completed sale; cancellations and returns are not present in the export.")
bullet(doc, "Revenue is recorded in Indian Rupees (₹) and is inclusive of the unit price only — no tax column exists.")
bullet(doc, "Missing City values are missing at random and can reasonably be inferred from the product category's typical city distribution.")
bullet(doc, "Duplicate Order_ID rows are accidental duplicates, not legitimate split shipments.")
bullet(doc, "The dataset spans January 2025 to May 2026 and is treated as a representative sample of trading activity.")

heading(doc, "10.2 Constraints and Limitations", 2)
bullet(doc, "The dataset is synthetic; absolute revenue figures do not represent a real company.")
bullet(doc, "`Category` is not a fixed product attribute in this dataset — every product appears under both "
            "Accessories and Electronics, and every product sells across the full price range. Category "
            "revenue shares are reported as transaction labels only; no product-line or price-positioning "
            "conclusion is drawn from them.")
bullet(doc, "Imputing 12 city values slightly smooths city-level results and is documented rather than hidden.")
bullet(doc, "Only about 17 months of history are available, so seasonality conclusions (for example the June–July dip) should be re-validated with more data.")
bullet(doc, "Outliers were retained, not removed. Their 41.5% revenue share is reported explicitly so the reader can judge the impact.")
bullet(doc, "The analysis is descriptive. No causal claim is made between payment mode, city or product and revenue.")

# ---------------------------- 11. PLAN ---------------------------- #
heading(doc, "11. Project Plan", 1)
add_table(
    doc,
    ["Day", "Planned activity", "Deliverable at end of day"],
    [
        ["Day 1", "Read brief, select dataset, set up Python environment, load and profile the data",
         "Environment ready; structural profile"],
        ["Day 2", "Data cleaning: missing values, duplicates, date conversion, revenue validation",
         "Cleaned dataset + documented cleaning decisions"],
        ["Day 3", "Core analysis: totals, product/city/customer rankings, monthly trend, payment comparison",
         "KPI and ranking tables"],
        ["Day 4", "NumPy statistics, IQR outlier detection, category summary, weekday pattern",
         "Statistical profile and outlier report"],
        ["Day 5", "Write recommendations, finalise SRS, push repository, prepare viva notes",
         "SRS + notebook + GitHub link ready for LMS"],
    ],
    widths=[1.8, 9.4, 4.8],
    font_size=8.5,
    first_col_bold=True,
)

# ---------------------------- 12. CONCLUSION ---------------------------- #
heading(doc, "12. Conclusion", 1)
para(doc,
     "This SRS defines a complete, reproducible approach to the Retail Sales Performance Analysis "
     "project. Starting from a raw export containing missing values and duplicate orders, the "
     "pipeline produces a validated revenue figure, ranks products, cities and customers, tracks "
     "monthly performance, compares payment channels, isolates statistical outliers and summarises "
     "revenue dispersion with NumPy. Every requirement in Section 5 is traceable to an executed "
     "notebook cell, and every recommendation in the notebook is backed by a computed number rather "
     "than an opinion.")

# ---------------------------- APPENDIX A ---------------------------- #
doc.add_page_break()
heading(doc, "Appendix A — Result Snapshot", 1)
para(doc, "Headline results produced by the executed notebook (500 orders after cleaning):",
     italic=True, space_after=8)
add_table(
    doc,
    ["Metric", "Result"],
    [
        ["Total revenue", "₹ 1,39,54,120.00"],
        ["Total quantity sold", "1,480 units"],
        ["Unique orders", "500"],
        ["Average order value", "₹ 27,908.24"],
        ["Median order value", "₹ 4,998.00"],
        ["Largest single order", "₹ 2,29,995.00"],
        ["Top product by revenue", "Keyboard — ₹ 34,38,656 (24.6% share)"],
        ["Top city by revenue", "Delhi — ₹ 33,97,468 (24.3% share)"],
        ["Best month", "February — ₹ 20,03,009"],
        ["Top customer", "C0018 — ₹ 4,45,986 across 4 orders"],
        ["Leading payment mode", "UPI — ₹ 42,46,098 (30.4% share)"],
        ["IQR outliers", "32 orders above ₹ 93,897 (41.5% of revenue)"],
        ["Largest category", "Accessories — ₹ 74,92,779 (53.7% share)"],
        ["Revenue distribution", "Mean ₹ 27,908 · Median ₹ 4,998 · Std ₹ 47,747"],
    ],
    widths=[6.0, 10.0],
    font_size=9,
    first_col_bold=True,
)

# ---------------------------- APPENDIX B ---------------------------- #
heading(doc, "Appendix B — References", 1)
refs = [
    "Pandas Development Team. Pandas Documentation — User Guide (groupby, datetime, I/O). https://pandas.pydata.org/docs/",
    "NumPy Developers. NumPy Reference — Statistics and percentile functions. https://numpy.org/doc/stable/",
    "McKinney, W. (2022). Python for Data Analysis, 3rd Edition. O'Reilly Media.",
    "Project brief: Python, NumPy & Pandas — Individual Industry Projects, Project 1: Retail Sales Performance Analysis (course handout).",
    "Tukey, J. W. (1977). Exploratory Data Analysis. Addison-Wesley — source of the 1.5 × IQR rule.",
]
for r_text in refs:
    bullet(doc, r_text)

doc.save(OUT)
print("saved:", OUT)
print("size:", os.path.getsize(OUT), "bytes")
