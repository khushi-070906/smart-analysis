// Builds Supermarket_Sales_Analysis_Report.docx (Times New Roman, black text, no colours).
// Usage: node build_report.js   (run report_charts.py first to create charts_bw/)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel, BorderStyle,
  WidthType, PageNumber, PageBreak, VerticalAlign,
} = require("docx");

const ROOT = __dirname;
const OUT = path.join(ROOT, "Supermarket_Sales_Analysis_Report.docx");

// ---------- data ----------
const lines = fs.readFileSync(path.join(ROOT, "data/supermarket_sales_500_rows.csv"), "utf8").trim().split("\n");
const head = lines[0].split(",");
const rows = lines.slice(1).map(l => {
  const v = l.split(",");
  const r = {};
  head.forEach((h, i) => (r[h] = v[i]));
  r.Quantity = +r.Quantity; r["Unit Price"] = +r["Unit Price"]; r.Rating = +r.Rating; r.Sales = +r.Sales;
  return r;
});
const TOTAL = rows.reduce((a, r) => a + r.Sales, 0);
const sum = a => a.reduce((x, r) => x + r.Sales, 0);
function group(key) {
  const m = new Map();
  rows.forEach(r => { const k = typeof key === "function" ? key(r) : r[key]; if (!m.has(k)) m.set(k, []); m.get(k).push(r); });
  return [...m].map(([k, rs]) => ({
    k, rs, n: rs.length, sales: sum(rs), units: rs.reduce((a, r) => a + r.Quantity, 0),
    avg: sum(rs) / rs.length, rating: rs.reduce((a, r) => a + r.Rating, 0) / rs.length,
  }));
}
const inr = (n, d = 2) => "₹" + n.toLocaleString("en-IN", { minimumFractionDigits: d, maximumFractionDigits: d });
const pct = n => (n / TOTAL * 100).toFixed(1) + "%";
const bySales = (a, b) => b.sales - a.sales;

// ---------- layout helpers (all black, Times New Roman) ----------
const FONT = "Times New Roman";
const SIZE = 24; // 12 pt
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: SIZE, color: "000000", ...o });
const para = t => new Paragraph({ spacing: { after: 160, line: 320 }, alignment: AlignmentType.JUSTIFIED, children: [run(t)] });
const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run(t, { size: 30, bold: true })] });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run(t, { size: 26, bold: true })] });
const bullet = t => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 100, line: 300 }, children: [run(t)] });
const num = (ref, t) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 100, line: 300 }, children: [run(t)] });
const spacer = () => new Paragraph({ spacing: { after: 120 }, children: [] });

const line = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: line, bottom: line, left: line, right: line };
function table(headers, data, widths, aligns = []) {
  const mk = (txt, i, isHead, keep) =>
    new TableCell({
      width: { size: widths[i], type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
      margins: { top: 70, bottom: 70, left: 110, right: 110 },
      children: [new Paragraph({ alignment: aligns[i] || AlignmentType.LEFT, keepNext: keep,
        children: [run(String(txt), { size: 22, bold: isHead })] })],
    });
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) => mk(h, i, true, true)) }),
      ...data.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => mk(c, i, false, ri < data.length - 1)) })),
    ],
  });
}
function figure(file, caption) {
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 60 }, keepNext: true,
      children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(ROOT, "charts_bw", file)),
        transformation: { width: 520, height: 292 },
        altText: { title: caption, description: caption, name: file } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 },
      children: [run(caption, { italics: true, size: 22 })] }),
  ];
}
const R = AlignmentType.RIGHT, L = AlignmentType.LEFT;

// ---------- numbers ----------
const prod = group("Product").sort(bySales);
const branch = group("Branch").sort(bySales);
const cat = group("Category").sort(bySales);
const pay = group("Payment").sort((a, b) => b.n - a.n);
const cust = group("Customer Type").sort(bySales);
const monthOrder = ["01", "02", "03", "04", "05", "06", "07"];
const monthName = { "01": "January", "02": "February", "03": "March", "04": "April", "05": "May", "06": "June", "07": "July (1 day only)" };
const months = group(r => r.Date.slice(5, 7)).sort((a, b) => monthOrder.indexOf(a.k) - monthOrder.indexOf(b.k));
const cityOf = Object.fromEntries(rows.map(r => [r.Branch, r.City]));
const avgRating = rows.reduce((a, r) => a + r.Rating, 0) / rows.length;
const units = rows.reduce((a, r) => a + r.Quantity, 0);
const member = cust.find(c => c.k === "Member"), normal = cust.find(c => c.k === "Normal");

// ---------- content ----------
const c = [];

// Cover page
c.push(
  new Paragraph({ spacing: { before: 2800, after: 240 }, alignment: AlignmentType.CENTER,
    children: [run("DATA ANALYTICS PROJECT REPORT", { size: 26, bold: true })] }),
  new Paragraph({ spacing: { after: 200 }, alignment: AlignmentType.CENTER,
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "000000", space: 12 } },
    children: [run("Supermarket Sales Analysis", { size: 60, bold: true })] }),
  new Paragraph({ spacing: { before: 240, after: 1800 }, alignment: AlignmentType.CENTER,
    children: [run("Analysis of 500 supermarket sales transactions", { size: 28 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [run("Prepared by")] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [run("Khushi Mittal", { size: 30, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [run("October 2026")] }),
  new Paragraph({ children: [new PageBreak()] }),
);

// 1
c.push(
  h1("1. Problem Statement"),
  para("The aim of this project is to study supermarket sales data and find useful information about products, branches, categories, customers, payment methods and ratings."),
  para("A supermarket can use this information to decide what to stock, which branches need more attention, which payment options to support and how to plan membership offers. This report answers the following six questions:"),
  num("q", "Which product has the highest sales?"),
  num("q", "Which branch performs best?"),
  num("q", "Which category sells the most?"),
  num("q", "What is the most popular payment method?"),
  num("q", "Do Members spend more than Normal customers?"),
  num("q", "What is the average customer rating?"),
  spacer(),
);

// 2
c.push(
  h1("2. Dataset"),
  para(`The dataset has 500 sales transactions from 1 January 2026 to 1 July 2026. The sales come from four branches, and each branch is in a different city: A (Jaipur), B (Delhi), C (Mumbai) and D (Bengaluru). There are 20 products in 8 categories. Together, the transactions sold ${units.toLocaleString("en-IN")} units and earned ${inr(TOTAL)}, which is an average of ${inr(TOTAL / rows.length)} per transaction.`),
  para("The dataset has the following columns:"),
  table(["Column", "Description"], [
    ["Invoice ID", "Unique number of the transaction (INV0001 to INV0500)"],
    ["Date", "Date of the sale"],
    ["Branch, City", "Store branch (A to D) and the city it is in"],
    ["Customer Type", "Member or Normal"],
    ["Gender", "Male or Female"],
    ["Product, Category", "Item sold and the category it belongs to"],
    ["Quantity", "Number of units sold (1 to 10)"],
    ["Unit Price", "Price of one unit in rupees"],
    ["Payment", "UPI, Net Banking, Card or Cash"],
    ["Rating", "Customer rating (3.0 to 5.0)"],
    ["Sales", "Quantity × Unit Price"],
  ], [2300, 6726]),
  spacer(),
  h2("Checking the data"),
  bullet("There are no missing values."),
  bullet("All 500 invoice IDs are different, so there are no duplicate transactions."),
  bullet("Sales is equal to Quantity × Unit Price in every row."),
  bullet("Quantities are between 1 and 10 and all prices are above zero, so no rows had to be removed."),
  spacer(),
);

// 3
c.push(
  h1("3. Methodology"),
  para("The analysis was done in the following steps:"),
  num("m", "Load the CSV dataset into Python using pandas."),
  num("m", "Check the data for missing values, duplicate invoices and wrong quantities or prices."),
  num("m", "Calculate Sales = Quantity × Unit Price for each transaction."),
  num("m", "Group the data by product, branch, category, payment method, customer type and month, and find totals, counts and averages."),
  num("m", "Draw charts to compare the results. A simple dashboard was also made as an HTML page, with filters for branch and customer type."),
  num("m", "Use the results to suggest business decisions."),
  para("Python with the pandas and matplotlib libraries was used for the analysis. The totals in this report match the totals shown in the dashboard."),
  spacer(),
);

// 4
const top10 = prod.slice(0, 10);
c.push(
  h1("4. Findings"),
  h2("4.1 Product sales"),
  para(`Cheese has the highest sales at ${inr(prod[0].sales)}, which is ${pct(prod[0].sales)} of total sales. Coffee (${inr(prod[1].sales)}) and Shampoo (${inr(prod[2].sales)}) are very close behind. These three products are within 1.5% of each other and together make up ${pct(prod[0].sales + prod[1].sales + prod[2].sales)} of all sales.`),
  para("The ranking depends on price more than on the number of units sold. Cheese sold only 126 units, but each unit costs about ₹220. Potato sold 194 units but earned only ₹6,799.38. So costly items add more to sales than items that are bought often."),
  table(["Rank", "Product", "Units", "Sales", "Share"],
    top10.map((p, i) => [i + 1, p.k, p.units, inr(p.sales), pct(p.sales)]),
    [900, 2800, 1300, 2326, 1700], [R, L, R, R, R]),
  ...figure("top_products.png", "Figure 1. Top 10 products by sales"),

  h2("4.2 Branch performance"),
  para(`Branch C (Mumbai) performs best, with sales of ${inr(branch[0].sales)}, which is ${pct(branch[0].sales)} of total sales. It has the most transactions (143, which is 36% more than Branch A in Jaipur) and also the highest average transaction at ${inr(branch[0].avg)}.`),
  para("Branch A (Jaipur) has the lowest sales and the lowest average rating (3.84). Branch D (Bengaluru) has the best rating (4.09), but its average transaction is the smallest, so customers there buy smaller amounts."),
  table(["Branch", "Sales", "Share", "Transactions", "Avg. transaction", "Avg. rating"],
    branch.map(b => [`${b.k} (${cityOf[b.k]})`, inr(b.sales), pct(b.sales), b.n, inr(b.avg), b.rating.toFixed(2)]),
    [1900, 1700, 1000, 1500, 1726, 1200], [L, R, R, R, R, R]),
  ...figure("branch_sales.png", "Figure 2. Sales by branch"),

  h2("4.3 Category sales"),
  para(`Beverages sells the most, with sales of ${inr(cat[0].sales)} (${pct(cat[0].sales)} of the total). Beverages, Personal Care and Dairy together make up ${pct(cat[0].sales + cat[1].sales + cat[2].sales)} of all sales.`),
  para("Snacks had 75 transactions, which is the third highest, but only ₹16,992.97 in sales because snack items are cheap. Bakery is the smallest category with ₹6,516.10."),
  table(["Category", "Sales", "Share", "Transactions"],
    cat.map(x => [x.k, inr(x.sales), pct(x.sales), x.n]),
    [3000, 2326, 1700, 2000], [L, R, R, R]),
  ...figure("category_sales.png", "Figure 3. Sales by category"),

  h2("4.4 Payment methods"),
  para(`UPI is the most used payment method with ${pay[0].n} transactions (${(pay[0].n / 5).toFixed(1)}%). However, it leads Net Banking by only one transaction. All four methods are between 122 and 127 transactions, so customers use all of them almost equally. UPI, Net Banking and Card together account for 75.6% of the transactions.`),
  table(["Payment method", "Transactions", "Share of transactions", "Sales"],
    pay.map(p => [p.k, p.n, (p.n / 5).toFixed(1) + "%", inr(p.sales)]),
    [2800, 1800, 2426, 2000], [L, R, R, R]),
  ...figure("payment_methods.png", "Figure 4. Transactions by payment method"),

  h2("4.5 Members and Normal customers"),
  para(`Members do not spend more than Normal customers. The average Member transaction is ${inr(member.avg)}, while Normal customers average ${inr(normal.avg)}. This is a difference of ₹13.93 per transaction, or 2.9%.`),
  para("Members still bring in more money overall because they shop more often. They made 296 of the 500 transactions (59.2%) and 58.5% of total sales. So their value comes from how often they buy, not from how much they spend each time. Men and women spend almost the same amount per transaction (₹491.66 and ₹486.09)."),
  table(["Customer type", "Transactions", "Sales", "Share of sales", "Avg. transaction", "Avg. rating"],
    cust.map(x => [x.k, x.n, inr(x.sales), pct(x.sales), inr(x.avg), x.rating.toFixed(2)]),
    [1600, 1550, 1700, 1350, 1600, 1226], [L, R, R, R, R, R]),
  ...figure("customer_type.png", "Figure 5. Average transaction of Members and Normal customers"),

  h2("4.6 Customer ratings"),
  para(`The average customer rating is ${avgRating.toFixed(2)} out of 5, and the middle (median) rating is 4.0. About 53.2% of the transactions were rated 4.0 or higher. All ratings in the data are between 3.0 and 5.0, so there are no very low ratings.`),
  para("Ratings are similar across branches (3.84 to 4.09) and categories. Fruits has the lowest category rating (3.83) and Bakery has the highest (4.24)."),

  h2("4.7 Sales by month"),
  para("Sales were highest in April at ₹52,569.77 from 90 transactions. February was the weakest full month with ₹30,068.15 from 65 transactions. July has data only for 1 July, so its low value does not mean that sales fell."),
  table(["Month (2026)", "Sales", "Transactions", "Avg. transaction"],
    months.map(m => [monthName[m.k], inr(m.sales), m.n, inr(m.avg)]),
    [3000, 2326, 1700, 2000], [L, R, R, R]),
  ...figure("monthly_trend.png", "Figure 6. Sales by month (July has 1 day of data only)"),
);

// 5
c.push(
  h1("5. Summary of Results"),
  table(["Question", "Answer"], [
    ["Highest-selling product", `Cheese, ${inr(prod[0].sales)}`],
    ["Best branch", `C (Mumbai), ${inr(branch[0].sales)}`],
    ["Category that sells the most", `Beverages, ${inr(cat[0].sales)}`],
    ["Most popular payment method", `UPI, ${pay[0].n} transactions (one more than Net Banking)`],
    ["Do Members spend more?", "No. Members average ₹483.14 and Normal customers average ₹497.07"],
    ["Average customer rating", `${avgRating.toFixed(2)} out of 5`],
    ["Total sales", `${inr(TOTAL)} from 500 transactions`],
  ], [3300, 5726]),
  spacer(),
);

// 6
c.push(
  h1("6. Business Decisions"),
  num("d", "Keep more stock of Cheese, Coffee and Shampoo, and of the Beverages, Personal Care and Dairy categories. Together these categories give about 60% of sales."),
  num("d", "Find out why Branch C (Mumbai) does well and try the same practices in the other branches. Start with Jaipur, which has the fewest transactions and the lowest rating."),
  num("d", "Keep all four payment methods. They are used almost equally, so none should be removed. UPI should be kept running smoothly because it is used the most."),
  num("d", "Improve the membership offers. Members buy more often but spend ₹13.93 less per transaction than Normal customers. Offers such as bundle discounts or discounts on a minimum bill could encourage bigger purchases."),
  num("d", "Work on customer ratings. Check Jaipur (3.84) and the Fruits category (3.83), which have the lowest ratings, for problems with service, freshness and stock."),
  num("d", "Study the sales pattern over the months, such as the low sales in February and the high sales in April, before planning stock or offers. More than one year of data is needed to be sure."),
  spacer(),
);

// 7
c.push(
  h1("7. Limitations"),
  bullet("The data covers only 1 January to 1 July 2026, so July is incomplete and sales cannot be compared with earlier years."),
  bullet("There is no cost or profit data, so “best” here means the highest sales, not the highest profit."),
  bullet("Each branch is in one city, so the effect of the branch cannot be separated from the effect of the city."),
  bullet("There is no customer ID, so repeat purchases cannot be measured."),
  bullet("The payment methods are used almost equally, so their ranking is not a strong result."),
  spacer(),
);

// 8
c.push(
  h1("8. Conclusion"),
  para(`This project turned 500 sales records into simple answers. Cheese is the top product, Branch C (Mumbai) is the best branch, Beverages is the top category and the average rating is ${avgRating.toFixed(2)} out of 5. Two results were different from what one might expect: Members spend slightly less per transaction than Normal customers, and all payment methods are used almost equally. The supermarket can use these results to stock its best products, copy the practices of its best branch and plan better membership offers.`),
);

// ---------- document ----------
const indent = { paragraph: { indent: { left: 720, hanging: 360 } } };
const doc = new Document({
  creator: "Khushi Mittal",
  title: "Supermarket Sales Analysis - Project Report",
  styles: {
    default: { document: { run: { font: FONT, size: SIZE, color: "000000" } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: FONT, color: "000000" },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: "000000" },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: indent }] },
      ...["q", "m", "d"].map(ref => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: indent }] })),
    ],
  },
  sections: [{
    properties: { titlePage: true,
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: {
      default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000", space: 4 } },
        children: [run("Supermarket Sales Analysis", { size: 20 })] })] }),
      first: new Header({ children: [new Paragraph({ children: [] })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [run("Page ", { size: 20 }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20, color: "000000" })] })] }),
      first: new Footer({ children: [new Paragraph({ children: [] })] }),
    },
    children: c,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("wrote", OUT, buf.length, "bytes");
});
