"""
Black-and-grey versions of the six report charts (Times-style font, value labels).
Usage: python report_charts.py data/supermarket_sales_500_rows.csv
Writes charts_bw/*.png (used by build_report.js).
"""
import os
import sys
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.family"] = "serif"
plt.rcParams["font.serif"] = ["Times New Roman", "Liberation Serif", "DejaVu Serif"]
plt.rcParams["font.size"] = 11

path = sys.argv[1] if len(sys.argv) > 1 else "data/supermarket_sales_500_rows.csv"
df = pd.read_csv(path)
df["Sales"] = df["Quantity"] * df["Unit Price"]
os.makedirs("charts_bw", exist_ok=True)


def bar(series, title, fname, ylabel, fmt="{:,.0f}"):
    fig, ax = plt.subplots(figsize=(8, 4.5))
    bars = ax.bar([str(i) for i in series.index], series.values, color="#7f7f7f", edgecolor="black", linewidth=0.8)
    for b, v in zip(bars, series.values):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height(), fmt.format(v),
                ha="center", va="bottom", fontsize=9)
    ax.set_title(title, fontweight="bold")
    ax.set_ylabel(ylabel)
    ax.set_ylim(0, series.max() * 1.12)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    plt.xticks(rotation=40, ha="right")
    plt.tight_layout()
    plt.savefig(f"charts_bw/{fname}.png", dpi=150)
    plt.close()


prod = df.groupby("Product")["Sales"].sum().sort_values(ascending=False).head(10)
bar(prod, "Top 10 Products by Sales", "top_products", "Sales (Rs)")

br = df.groupby(["Branch", "City"])["Sales"].sum().sort_values(ascending=False)
br.index = [f"{b} ({c})" for b, c in br.index]
bar(br, "Sales by Branch", "branch_sales", "Sales (Rs)")

cat = df.groupby("Category")["Sales"].sum().sort_values(ascending=False)
bar(cat, "Sales by Category", "category_sales", "Sales (Rs)")

pay = df["Payment"].value_counts()
bar(pay, "Transactions by Payment Method", "payment_methods", "Transactions")

cust = df.groupby("Customer Type")["Sales"].mean().sort_values(ascending=False)
bar(cust, "Average Transaction: Member vs Normal", "customer_type", "Average sales (Rs)", "{:,.2f}")

m = df.assign(M=pd.to_datetime(df["Date"]).dt.to_period("M")).groupby("M")["Sales"].sum()
m.index = [p.strftime("%b %Y") + (" (1 day)" if str(p) == "2026-07" else "") for p in m.index]
bar(m, "Monthly Sales", "monthly_trend", "Sales (Rs)")
print("charts_bw/ written")
