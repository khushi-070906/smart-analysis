"""
Supermarket Sales Analysis
Usage:  python supermarket_analysis.py supermarket_sales_500_rows.csv
Outputs: printed answers, charts/*.png, results.md
"""
import sys
import os
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

path = sys.argv[1] if len(sys.argv) > 1 else "supermarket_sales_500_rows.csv"
df = pd.read_csv(path) if path.endswith(".csv") else pd.read_excel(path)


def find(*keys):
    """Find a column by loose name match."""
    for c in df.columns:
        n = c.lower().replace("_", " ").strip()
        if any(k in n for k in keys):
            return c
    raise KeyError(f"No column matching {keys}. Columns: {list(df.columns)}")


C = {
    "inv": find("invoice"),
    "date": find("date"),
    "branch": find("branch"),
    "city": find("city"),
    "cust": find("customer"),
    "prod": find("product"),
    "cat": find("category"),
    "qty": find("quantity", "qty"),
    "price": find("unit price", "price"),
    "pay": find("payment"),
    "rating": find("rating"),
}

# ---- 1. Clean / validate ----
print("Rows:", len(df))
print("Missing values per column:\n", df.isna().sum()[df.isna().sum() > 0], "\n")
df = df.drop_duplicates(subset=C["inv"])
df[C["date"]] = pd.to_datetime(df[C["date"]], errors="coerce")
df = df.dropna(subset=[C["qty"], C["price"]])
df = df[(df[C["qty"]] > 0) & (df[C["price"]] > 0)]

# ---- 2. Sales = Quantity x Unit Price ----
df["Sales"] = df[C["qty"]] * df[C["price"]]

# ---- 3. Answers ----
by_prod = df.groupby(C["prod"])["Sales"].sum().sort_values(ascending=False)
by_branch = df.groupby([C["branch"], C["city"]])["Sales"].sum().sort_values(ascending=False)
by_cat = df.groupby(C["cat"])["Sales"].sum().sort_values(ascending=False)
by_pay = df[C["pay"]].value_counts()
by_cust = df.groupby(C["cust"])["Sales"].mean()
avg_rating = df[C["rating"]].mean()

(b_branch, b_city), b_sales = next(iter(by_branch.items()))
lines = [
    "# Supermarket Sales Analysis - Results",
    f"- Transactions analysed: {len(df)}",
    f"- Total sales: Rs {df['Sales'].sum():,.2f}",
    f"- Highest-selling product: {by_prod.index[0]} (Rs {by_prod.iloc[0]:,.2f})",
    f"- Best branch: {b_branch} ({b_city}) with Rs {b_sales:,.2f}",
    f"- Top category: {by_cat.index[0]} (Rs {by_cat.iloc[0]:,.2f})",
    f"- Most popular payment method: {by_pay.index[0]} ({by_pay.iloc[0]} transactions)",
    "- Average transaction by customer type: "
    + ", ".join(f"{k} Rs {v:,.2f}" for k, v in by_cust.items()),
    f"- Average rating: {avg_rating:.2f} / 5",
]
text = "\n".join(lines)
print(text)
open("results.md", "w", encoding="utf-8").write(text + "\n")

# ---- 4. Charts ----
os.makedirs("charts", exist_ok=True)


def bar(series, title, fname, ylabel="Sales (Rs)"):
    ax = series.plot(kind="bar", figsize=(8, 4.5), color="#2f6fed")
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.set_xlabel("")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    plt.savefig(f"charts/{fname}.png", dpi=150)
    plt.close()


bar(by_prod.head(10), "Top 10 Products by Sales", "top_products")
branch_plot = by_branch.copy()
branch_plot.index = [f"{b} ({c})" for b, c in branch_plot.index]
bar(branch_plot, "Sales by Branch", "branch_sales")
bar(by_cat, "Sales by Category", "category_sales")
bar(by_pay, "Transactions by Payment Method", "payment_methods", "Transactions")
bar(by_cust, "Average Transaction: Member vs Normal", "customer_type", "Avg Sales (Rs)")

monthly = df.set_index(C["date"]).resample("MS")["Sales"].sum()
monthly.index = monthly.index.strftime("%b %Y")
bar(monthly, "Monthly Sales Trend", "monthly_trend")
print("\nCharts saved in ./charts, summary in results.md")
