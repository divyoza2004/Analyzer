# 📊 Sales Analyzer: Excel Data Analysis Workbook

A 10-task Excel project that turns a raw sales table (200 orders) into rankings, what-if models, statistics, pivots, charts and a one-page dashboard, using only built-in Excel features.

---

## 🎯 Objective

| # | Objective |
|---|-----------|
| 1 | Clean, structure and summarise a raw sales dataset of **200 orders** from **30 customers**, **5 regions** and **5 product categories** (Apr 2024 – Mar 2025). |
| 2 | Identify the **best customers** and the **high-value tier** using ranking, `INDEX`/`MATCH` and filters. |
| 3 | Test how **discount changes affect profit** (What-If, sensitivity table, scenarios). |
| 4 | Apply **statistics**: regression and descriptive statistics with a histogram. |
| 5 | Track **monthly growth** with up/down arrows and add a **timestamp** column. |
| 6 | Summarise results with a **pivot-style table, charts and a dashboard**. |

---

## ▶ Demo Video

<a href="https://drive.google.com/file/d/1nbVZaQ3YuRXYZnbgz7qX1-A7oWOnN7sA/view?usp=sharing" target="_blank" rel="noopener noreferrer">
  <img src="https://img.shields.io/badge/▶-Watch%20Demo%20Video-181717?style=for-the-badge&logo=github&logoColor=white" alt="Watch Demo Video" />
</a>

## ✨ Features

| Feature | What it does | Where |
|---------|--------------|-------|
| Top-10 customer highlighting | Ranks customers by total purchase and highlights the top 10 | `Q1_TopCustomers` |
| Data bars | In-cell bars show purchase size | `Q1_TopCustomers` |
| What-If model | Change the discount and uplift inputs to see new sales, cost and profit | `Q2_WhatIf` |
| Sensitivity table | Profit at 0%–25% discount side by side | `Q2_WhatIf` |
| Scenario comparison | Low, Current-like and Aggressive discount scenarios | `Q2_WhatIf` |
| Linear regression | Profit vs Sales equation, R², standard error, prediction | `Q3_Regression` |
| Descriptive statistics | 20+ statistics for Sales, Quantity, Discount, Profit | `Q4_DescStats` |
| Histogram | Sales distribution in 200-wide bins | `Q4_DescStats` |
| Growth arrows | ▲ ▼ ► on month-over-month sales growth, coloured green/red | `Q5_GrowthArrows` |
| Frozen timestamp | Time stamps once and stays fixed | `Q6_Timestamp` |
| Region filter | Dropdown (All / Central / East / North / South / West) | `Q7_HighValue` |
| Tiering | Each customer is labelled High-Value, Mid or Low | `Q7_HighValue` |
| Pivot summary | Sales, Profit and Profit Margin by Region × Category with colour scales | `Q8_PivotRegionProduct` |
| Charts | Bar, line and pie charts for the main KPIs | `Q9_Charts` |
| KPI dashboard | 6 KPI cards, 3 charts and a Top-5 customers list | `Q10_Dashboard` |

---

## 🧰 Requirements

| Type | Requirement |
|------|-------------|
| Software | Microsoft Excel 2016 or newer (Excel 365 recommended). LibreOffice Calc also opens it. |
| File | `Analyzer.xlsx` (11 sheets: `Raw Data` + `Q1` to `Q10`) |
| Setting | **Iterative calculation must be ON** (File ▸ Options ▸ Formulas ▸ Enable iterative calculation). Required for the frozen timestamp in Q6. |
| Add-in | **Analysis ToolPak** is only needed if you want to redo Q3 and Q4 through the ToolPak dialogs. The sheets already use formulas. |
| Input data | `Raw Data` sheet, columns A to I (see table below) |

### Input data columns (`Raw Data`)

| Column | Field | Type | Meaning |
|--------|-------|------|---------|
| A | Customer_ID | Text | Customer code, `CUST001` to `CUST030` |
| B | Customer_Name | Text | Name on the order |
| C | Region | Text | Central, East, North, South, West |
| D | Product_Category | Text | Books, Clothing, Electronics, Furniture, Office Supplies |
| E | Sales | Number | Order value after discount |
| F | Quantity | Number | Units sold (1 to 19) |
| G | Discount | Number | Discount rate (0 to 0.20) |
| H | Order_Date | Date | Order date |
| I | Profit | Number | Profit on the order |

> ⚠️ **Workbook note (Q2):** the *Profit vs Current* column (I6:I11) points one row too high (`=H5-$B$4` sits beside the 5% row but uses the 0% profit), so each value appears one row below its discount. Fix it by moving the formulas up one row (I5 `=H5-$B$4`, I6 `=H6-$B$4`, and so on).

> ⚠️ **Data note:** one `Customer_ID` can appear with different names (for example `CUST005` is both *Robert Anderson* and *Sarah Anderson*). The workbook therefore reports the **first name seen** for each ID.

---

## 🗂️ Task Explanations

### Task summary

| Task | Sheet | Title | Main Excel tools |
|------|-------|-------|------------------|
| Q1 | `Q1_TopCustomers` | Highlight top 10 customers | `SUMIF`, `COUNTIF`, `RANK`, `LARGE`, Conditional Formatting |
| Q2 | `Q2_WhatIf` | What-If discount analysis | `SUMPRODUCT`, sensitivity table, Scenario Manager |
| Q3 | `Q3_Regression` | Linear regression (Profit vs Sales) | `SLOPE`, `INTERCEPT`, `RSQ`, `LINEST`, scatter chart |
| Q4 | `Q4_DescStats` | Descriptive statistics + histogram | `AVERAGE`, `MEDIAN`, `STDEV`, `KURT`, `SKEW`, `QUARTILE`, `CONFIDENCE` |
| Q5 | `Q5_GrowthArrows` | Growth arrows | `SUMIFS`, `EOMONTH`, `IF`, Icon Sets |
| Q6 | `Q6_Timestamp` | Timestamp column | `NOW`, `IF`, iterative calculation |
| Q7 | `Q7_HighValue` | High-value customers | `INDEX`, `MATCH`, `PERCENTILE`, filter, data validation |
| Q8 | `Q8_PivotRegionProduct` | Pivot of Region × Product | `SUMIFS`, colour scales |
| Q9 | `Q9_Charts` | Bar, line, pie charts | `SUMIF`, `SUMIFS`, charts |
| Q10 | `Q10_Dashboard` | KPI dashboard | `SUM`, `AVERAGE`, `INDEX`/`MATCH`, charts, data bars |

### Detailed explanation

| Task | What the task asks | How it is solved | Key result |
|------|--------------------|------------------|------------|
| **Q1** Top customers | Highlight the top 10 customers by total purchase. | `SUMIF` totals each customer's sales, `COUNTIF` counts orders, `RANK` orders them. A conditional-format rule `=$E3<=10` highlights those rows, and data bars show purchase size. `LARGE(...,10)` gives the cut-off. | Top customer is **CUST021** (13,074.26). The top 10 account for about **49%** of all purchases. 10th-place cut-off is 7,489.85. |
| **Q2** What-If | Show how changing the discount changes total profit. | List price is rebuilt as `Sales ÷ (1 − Discount)`. Volume change = `(new discount − current avg discount) × 100 × uplift per point`. New Sales = list price × (1 − discount) × (1 + volume change). New Cost = cost × (1 + volume change). Part B is a sensitivity table (0%–25%) and Part C compares 3 scenarios. | Current profit is **63,630.03** at an average discount of **10.5%**. A 10% discount with 2% uplift gives **64,090.75** (+460.72). Lower discounts earn more profit (0% gives 68,566.81), and from 15% upward profit falls below the current level. |
| **Q3** Regression | Fit a line of Profit on Sales. | `SLOPE`, `INTERCEPT`, `RSQ`, `CORREL`, `STEYX` and `LINEST` reproduce the Analysis ToolPak output. A scatter chart shows the fit. | **Profit = 14.64 + 0.3075 × Sales**, R² = **0.579**, 200 observations. Sales of 1,000 predicts profit of about **322.10**. |
| **Q4** Descriptive stats | Summarise Sales, Quantity, Discount and Profit, and draw a histogram. | Each statistic is a formula over the `Raw Data` column: mean, standard error, median, mode, SD, variance, kurtosis, skewness, range, min, max, sum, count, quartiles, IQR, coefficient of variation, 90th percentile, 95% confidence level. `COUNTIFS` builds the histogram bins. | Mean Sales **987.14**, Mean Profit **318.15**. Sales is almost symmetric (skewness 0.04). Most orders fall between 700 and 1,300. |
| **Q5** Growth arrows | Show ▲/▼ on monthly sales growth. | `SUMIFS` with `EOMONTH` totals each month. Growth % = this month ÷ last month − 1. The arrow is `=IF(D>0,"▲",IF(D<0,"▼","►"))`, coloured green or red by conditional formatting. | Best month: **Jul → Aug 2024 (+70.7%)**. Worst: **Jun → Jul 2024 (−42.4%)**. |
| **Q6** Timestamp | Add a timestamp column using `NOW()`. | `=IF(A4<>"",IF(C4="",NOW(),C4),"")` is a deliberate circular reference. With iterative calculation on, it stamps the time once when an ID is typed and then keeps it frozen. A second block shows the live `NOW()` stamp for every order. | Typing a Customer_ID in column A stamps the date and time in column C. |
| **Q7** High-value customers | Find high-value customers with `INDEX`, `MATCH` and filters. | `SUMIFS` totals respect a **Region dropdown**. The threshold is the 80th percentile (`PERCENTILE`). `IF` labels each customer High-Value, Mid or Low. A tie-breaker key (`total + ROW()/1,000,000`) lets `LARGE` + `INDEX`/`MATCH` build a Top-10 list. An AutoFilter covers A6:E36. | Threshold is **8,413.85**. **6 customers** are High-Value and they hold about **33%** of purchases. |
| **Q8** Pivot table | Show total sales by region and product. | Three `SUMIFS` grids stand in for pivot tables: (A) Sales, (B) Profit, (C) Profit margin (Profit ÷ Sales), each with totals and colour scales. | Largest cell: **Central × Books 12,717.92**. Best region by sales: **Central 44,968.63**. Overall margin **32.2%**. |
| **Q9** Charts | Make bar, line and pie charts for the KPIs. | Three formula tables feed three charts: **bar** = sales and profit by region, **line** = monthly sales and profit trend, **pie** = sales share by product category. | Books is the largest category (46,337.01). Monthly sales dip in Jul 2024 and again in Feb 2025. |
| **Q10** Dashboard | Summarise the insights on one page. | Six KPI cards (Total Sales, Total Profit, Margin, Orders, Avg Discount, Customers), a region table, a category table and a Top-5 customers list (`INDEX`/`MATCH` on Q1 ranks) with data bars, plus bar, pie and line charts. | Sales **197,427.52**, Profit **63,630.03**, Margin **32.2%**, **200** orders, **30** customers, Avg discount **9.8%**. |

---

## 🚀 How to Use

| Step | Action |
|------|--------|
| 1 | Open `Analyzer.xlsx` in Excel and enable **iterative calculation** (see Requirements). |
| 2 | Review or edit the data on the `Raw Data` sheet. All other sheets update automatically. |
| 3 | In `Q2_WhatIf`, change the **INPUT** cells (new discount rate, volume uplift per point). |
| 4 | In `Q7_HighValue`, pick a region in cell **B3** to filter the tiers. |
| 5 | In `Q6_Timestamp`, type a Customer_ID in column A to stamp the time. |
| 6 | Open `Q10_Dashboard` for the one-page summary. |

---

## 🔑 Key Insights

| Area | Finding |
|------|---------|
| Customers | Top 10 customers drive roughly half of all purchases. |
| Discounts | Lower discounts give higher profit under the assumed volume uplift. |
| Relationship | Sales explains about 58% of the variation in profit. |
| Products | Books has the highest sales, Office Supplies the highest margin (34.1%). |
| Regions | Central and West lead on sales. North has the best margin (33.0%). |
| Trend | Monthly sales are uneven, with sharp dips in Jul 2024 and Feb 2025. |

---

## 📁 File Structure

| Sheet | Purpose |
|-------|---------|
| `Raw Data` | Source table (200 rows × 9 columns) |
| `Q1_TopCustomers` to `Q10_Dashboard` | One sheet per task |
