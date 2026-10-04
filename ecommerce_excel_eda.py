"""E-commerce Excel EDA for EXCEL_PROJECT (1).xlsx.

Run:
    pip install pandas matplotlib seaborn openpyxl
    python ecommerce_excel_eda.py

The script reads the workbook's 'data processing' sheet, creates charts in
the visualizations folder, and saves a cleaned analysis-ready CSV file.
"""

from pathlib import Path
import warnings

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
sns.set_theme(style="whitegrid", palette="Set2")

# -----------------------------------------------------------------------------
# 1. File paths and data loading
# -----------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
FILE_PATH = BASE_DIR / "upload" / "sales_data.xlsx"
SHEET_NAME = "data processing"  # Already includes Month, Age Group and State.
OUTPUT_DIR = BASE_DIR / "visualizations"
OUTPUT_DIR.mkdir(exist_ok=True)

if not FILE_PATH.exists():
    raise FileNotFoundError(
        f"Workbook not found: {FILE_PATH}\n"
        "Update FILE_PATH at the top of this script before running it."
    )

df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

# -----------------------------------------------------------------------------
# 2. Standardise column names and data types
# -----------------------------------------------------------------------------
df.columns = (
    df.columns.astype(str)
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
    .str.replace("-", "_", regex=False)
)

df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df["Age"] = pd.to_numeric(df["Age"], errors="coerce")
df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce")
df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")

for column in ["Gender", "Status", "Channel", "Category", "Size", "STATE", "B2B"]:
    if column in df.columns:
        df[column] = df[column].astype("string").str.strip()

# Correct known spelling variations before calculations and visualizations.
df["Channel"] = df["Channel"].replace({
    "Menyntra": "Myntra",
    "AMenazon": "Amazon",
    "Meneesho": "Meesho",
})
df["Category"] = df["Category"].replace({
    "Womenestern Dress": "Western Dress",
    "kurta": "Kurta",
})

# Make Month a correctly ordered categorical variable for charts.
month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
if "MONTH" in df.columns:
    df["MONTH"] = df["MONTH"].astype("string").str[:3].str.title()
    df["MONTH"] = pd.Categorical(df["MONTH"], categories=month_order, ordered=True)

# -----------------------------------------------------------------------------
# 3. Initial data-quality checks
# -----------------------------------------------------------------------------
print("\n" + "=" * 70)
print("1. DATASET OVERVIEW")
print("=" * 70)
print(f"Rows: {df.shape[0]:,}")
print(f"Columns: {df.shape[1]}")
print("\nColumns:")
print(df.columns.tolist())
print("\nData types:")
print(df.dtypes)
print("\nFirst five rows:")
print(df.head())

print("\nMissing values:")
missing = df.isna().sum().sort_values(ascending=False)
print(missing[missing > 0] if (missing > 0).any() else "No missing values found.")

print(f"\nExact duplicate rows: {df.duplicated().sum():,}")
print(f"Duplicate Order IDs: {df['Order ID'].duplicated().sum():,}")
print("\nNumeric summary:")
print(df[["Age", "Qty", "Amount"]].describe().round(2))

# -----------------------------------------------------------------------------
# 4. Cleaning rules
# -----------------------------------------------------------------------------
# Keep an original copy before applying cleaning rules.
raw_df = df.copy()

# Remove exact duplicates, then keep valid order rows only.
df = df.drop_duplicates().copy()
df = df.dropna(subset=["Order ID", "Date", "Amount", "Qty"])
df = df[(df["Amount"] >= 0) & (df["Qty"] > 0)]

# Standardise common text fields.
for column in ["Gender", "Status", "Channel", "Category", "Size", "STATE", "B2B"]:
    if column in df.columns:
        df[column] = df[column].fillna("Unknown")

print("\n" + "=" * 70)
print("2. CLEANING SUMMARY")
print("=" * 70)
print(f"Rows before cleaning: {len(raw_df):,}")
print(f"Rows after cleaning:  {len(df):,}")
print(f"Rows removed:         {len(raw_df) - len(df):,}")

# -----------------------------------------------------------------------------
# 5. Feature engineering for EDA
# -----------------------------------------------------------------------------
df["Year"] = df["Date"].dt.year
df["Month Number"] = df["Date"].dt.month
df["Month Name"] = df["Date"].dt.month_name()
df["Weekday"] = df["Date"].dt.day_name()

if "MONTH" not in df.columns or df["MONTH"].isna().all():
    df["MONTH"] = pd.Categorical(
        df["Date"].dt.strftime("%b"), categories=month_order, ordered=True
    )

if "Age Group" not in df.columns:
    df["Age Group"] = pd.cut(
        df["Age"], bins=[0, 18, 25, 35, 45, 55, 65, np.inf],
        labels=["0-18", "19-25", "26-35", "36-45", "46-55", "56-65", "65+"]
    )

# -----------------------------------------------------------------------------
# 6. Excel-style KPI summary and grouped tables
# -----------------------------------------------------------------------------
total_revenue = df["Amount"].sum()
total_orders = df["Order ID"].nunique()
total_customers = df["Cust ID"].nunique()
total_quantity = df["Qty"].sum()
avg_order_value = total_revenue / total_orders

kpis = pd.DataFrame({
    "Metric": ["Total Revenue", "Unique Orders", "Unique Customers", "Units Sold", "Average Order Value"],
    "Value": [total_revenue, total_orders, total_customers, total_quantity, avg_order_value],
})

print("\n" + "=" * 70)
print("3. BUSINESS KPIs")
print("=" * 70)
print(kpis.to_string(index=False, formatters={"Value": "{:,.2f}".format}))

monthly_summary = (
    df.groupby("MONTH", observed=False)
    .agg(Revenue=("Amount", "sum"), Orders=("Order ID", "nunique"), Quantity=("Qty", "sum"))
    .reset_index()
)
monthly_summary["Average Order Value"] = monthly_summary["Revenue"] / monthly_summary["Orders"]

category_summary = (
    df.groupby("Category", dropna=False)
    .agg(Revenue=("Amount", "sum"), Orders=("Order ID", "nunique"), Quantity=("Qty", "sum"))
    .sort_values("Revenue", ascending=False)
    .reset_index()
)

channel_summary = (
    df.groupby("Channel", dropna=False)
    .agg(Revenue=("Amount", "sum"), Orders=("Order ID", "nunique"), Quantity=("Qty", "sum"))
    .sort_values("Revenue", ascending=False)
    .reset_index()
)

state_summary = (
    df.groupby("STATE", dropna=False)
    .agg(Revenue=("Amount", "sum"), Orders=("Order ID", "nunique"), Quantity=("Qty", "sum"))
    .sort_values("Revenue", ascending=False)
    .reset_index()
)

customer_summary = (
    df.groupby("Cust ID", dropna=False)
    .agg(Revenue=("Amount", "sum"), Orders=("Order ID", "nunique"), Quantity=("Qty", "sum"))
    .sort_values("Revenue", ascending=False)
    .head(10)
    .reset_index()
)

print("\nTop categories by revenue:\n", category_summary.to_string(index=False))
print("\nTop channels by revenue:\n", channel_summary.to_string(index=False))
print("\nTop 10 states by revenue:\n", state_summary.head(10).to_string(index=False))

# Save analysis tables for Excel, Power BI, or SQL use.
analysis_path = BASE_DIR / "cleaned_ecommerce_sales.csv"
df.to_csv(analysis_path, index=False)

summary_dir = BASE_DIR / "eda_summary_tables"
summary_dir.mkdir(exist_ok=True)
for file_name, table in {
    "kpis.csv": kpis,
    "monthly_summary.csv": monthly_summary,
    "category_summary.csv": category_summary,
    "channel_summary.csv": channel_summary,
    "state_summary.csv": state_summary,
    "top_customers.csv": customer_summary,
}.items():
    table.to_csv(summary_dir / file_name, index=False)

# -----------------------------------------------------------------------------
# 7. Reusable chart helper
# -----------------------------------------------------------------------------
def finish_chart(filename: str, title: str, x_label: str = "", y_label: str = ""):
    plt.title(title, fontsize=14, weight="bold")
    plt.xlabel(x_label)
    plt.ylabel(y_label)
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / filename, dpi=180, bbox_inches="tight")
    plt.show()


# -----------------------------------------------------------------------------
# 8. Visual EDA
# -----------------------------------------------------------------------------
# Chart 1: Monthly revenue trend
plt.figure(figsize=(11, 5))
sns.lineplot(data=monthly_summary, x="MONTH", y="Revenue", marker="o", linewidth=2.5)
plt.xticks(rotation=30)
finish_chart("01_monthly_revenue_trend.png", "Monthly Revenue Trend", "Month", "Revenue")

# Chart 2: Monthly order volume
plt.figure(figsize=(11, 5))
sns.barplot(data=monthly_summary, x="MONTH", y="Orders", color="#4C78A8")
plt.xticks(rotation=30)
finish_chart("02_monthly_orders.png", "Monthly Order Volume", "Month", "Unique Orders")

# Chart 3: Category revenue
plt.figure(figsize=(8, 5))
sns.barplot(data=category_summary, x="Category", y="Revenue", hue="Category", legend=False)
plt.xticks(rotation=20)
finish_chart("03_category_revenue.png", "Revenue by Category", "Category", "Revenue")

# Chart 4: Sales channel contribution
plt.figure(figsize=(9, 5))
sns.barplot(data=channel_summary, x="Channel", y="Revenue", hue="Channel", legend=False)
plt.xticks(rotation=35, ha="right")
finish_chart("04_channel_revenue.png", "Revenue by Sales Channel", "Channel", "Revenue")

# Chart 5: Gender and age-group purchase pattern
gender_age = (
    df.groupby(["Age Group", "Gender"], observed=False)["Amount"].sum().reset_index()
)
plt.figure(figsize=(11, 5))
sns.barplot(data=gender_age, x="Age Group", y="Amount", hue="Gender")
plt.xticks(rotation=25)
finish_chart("05_age_gender_revenue.png", "Revenue by Age Group and Gender", "Age Group", "Revenue")

# Chart 6: Top states by revenue
top_states = state_summary.head(10).sort_values("Revenue")
plt.figure(figsize=(10, 6))
sns.barplot(data=top_states, y="STATE", x="Revenue", hue="STATE", legend=False)
finish_chart("06_top_states_revenue.png", "Top 10 States by Revenue", "Revenue", "State")

# Chart 7: Order status breakdown
status_summary = df.groupby("Status", dropna=False)["Order ID"].nunique().sort_values(ascending=False).reset_index(name="Orders")
plt.figure(figsize=(10, 5))
sns.barplot(data=status_summary, x="Status", y="Orders", hue="Status", legend=False)
plt.xticks(rotation=35, ha="right")
finish_chart("07_order_status.png", "Orders by Status", "Status", "Unique Orders")

# Chart 8: Quantity distribution and outliers
plt.figure(figsize=(9, 5))
sns.histplot(df["Qty"], bins=30, kde=True, color="#59A14F")
finish_chart("08_quantity_distribution.png", "Quantity Distribution", "Quantity", "Number of Orders")

plt.figure(figsize=(9, 3.8))
sns.boxplot(x=df["Amount"], color="#F28E2B")
finish_chart("09_amount_outliers.png", "Order Amount Distribution and Outliers", "Order Amount")

# Chart 9: B2B vs consumer revenue
b2b_summary = df.groupby("B2B", dropna=False)["Amount"].sum().reset_index(name="Revenue")
plt.figure(figsize=(6, 5))
sns.barplot(data=b2b_summary, x="B2B", y="Revenue", hue="B2B", legend=False)
finish_chart("10_b2b_revenue.png", "B2B versus Consumer Revenue", "B2B", "Revenue")

# Chart 10: Correlation heatmap
numeric_columns = df[["Age", "Qty", "Amount"]].dropna()
plt.figure(figsize=(6, 4.5))
sns.heatmap(numeric_columns.corr(), annot=True, cmap="YlGnBu", vmin=-1, vmax=1)
finish_chart("11_correlation_heatmap.png", "Numeric Feature Correlation")

# Chart 11: Pareto chart (categories responsible for revenue)
pareto = category_summary.copy()
pareto["Cumulative %"] = pareto["Revenue"].cumsum() / pareto["Revenue"].sum() * 100
fig, ax1 = plt.subplots(figsize=(9, 5))
sns.barplot(data=pareto, x="Category", y="Revenue", ax=ax1, color="#4C78A8")
ax2 = ax1.twinx()
ax2.plot(pareto["Category"], pareto["Cumulative %"], color="#E45756", marker="o")
ax2.set_ylabel("Cumulative Revenue %")
ax2.set_ylim(0, 110)
finish_chart("12_category_pareto.png", "Category Revenue Pareto Analysis", "Category", "Revenue")

# -----------------------------------------------------------------------------
# 9. Simple EDA insights generated directly from data
# -----------------------------------------------------------------------------
best_month = monthly_summary.loc[monthly_summary["Revenue"].idxmax()]
best_category = category_summary.iloc[0]
best_channel = channel_summary.iloc[0]
best_state = state_summary.iloc[0]

print("\n" + "=" * 70)
print("4. AUTOMATIC EDA INSIGHTS")
print("=" * 70)
print(f"Highest-revenue month: {best_month['MONTH']} ({best_month['Revenue']:,.2f})")
print(f"Highest-revenue category: {best_category['Category']} ({best_category['Revenue']:,.2f})")
print(f"Highest-revenue channel: {best_channel['Channel']} ({best_channel['Revenue']:,.2f})")
print(f"Highest-revenue state: {best_state['STATE']} ({best_state['Revenue']:,.2f})")
print(f"\nSaved cleaned data: {analysis_path.name}")
print(f"Saved summary tables: {summary_dir.name}")
print(f"Saved charts folder: {OUTPUT_DIR.name}")

# Optional follow-up analyses to practise:
# 1. Compare category performance by Month using pd.crosstab(df['MONTH'], df['Category']).
# 2. Find highest-value customers: group by Cust ID and sum Amount.
# 3. Calculate cancellation rate: status_summary / total_orders * 100.
# 4. Build the same KPIs and charts in Power BI using cleaned_ecommerce_sales.csv.
