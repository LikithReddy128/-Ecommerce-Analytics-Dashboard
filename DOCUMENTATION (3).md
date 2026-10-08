# 📘 Project Documentation – Ecommerce Excel Dashboard

**Project:** Ecommerce Excel Dashboard  
**Tool:** Microsoft Excel (Tables, PivotTables, PivotCharts, Slicers, Map chart)  
**Data:** 2,400 orders · 1 Jan 2025 – 29 Mar 2025 · California  
**Workbook:** `dashboard/ecommerce-blank_excel_project.xlsx`

---

## Contents
1. [Project Aim / Goal](#1-project-aim--goal)
2. [Dataset Description](#2-dataset-description)
3. [Cleaning Process](#3-cleaning-process)
4. [Analysis Approach (KPIs, Questions, Slicers)](#4-analysis-approach)
5. [Final Result](#5-final-result)
6. [Outcomes and Insights](#6-outcomes-and-insights)
7. [Validation Notes](#7-validation-notes)
8. [Limitations and Next Steps](#8-limitations-and-next-steps)

---

## 1. Project Aim / Goal

**Aim:** Build a single-page, interactive Excel dashboard that helps an ecommerce business understand its sales performance and customer behaviour.

**Objectives**
- Track five headline KPIs: Orders, Quantity, Amount, Average Rating and Average Days to Deliver.
- Answer eight business questions (trend, buying channel, order size, popular products, gender split, location, shipping time, satisfaction).
- Let users filter the whole report with slicers (Order Mode and Gender Value).
- Publish the work as a professional, reproducible GitHub repository.

---

## 2. Dataset Description

| Item | Detail |
|---|---|
| Source file | `ecommerce-blank_dataset.xlsx` (sheet `Data`) |
| Rows / columns | 2,400 rows × 15 columns |
| Excel table name | `orders` (range A1:O2401) |
| Period | 1 Jan 2025 – 29 Mar 2025 (weeks 1–13) |
| Geography | State: California only; 58 distinct counties |
| Products | 20 (e.g. T-Shirts, Jeans, Sneakers, Tank Tops, Bikinis, Shorts, Sundresses, Graphic Tees, Hoodies & Sweatshirts, Sandals, Jewelry …) |
| Order modes | 5: App, Website, Target.com, Partner App, Instagram |
| Ratings | Whole numbers 1–5 |

### Column dictionary

| # | Column | Type | Origin | Notes |
|---|---|---|---|---|
| 1 | TX ID | Text | Raw | Unique; 0 duplicates |
| 2 | Product | Text | Raw | 20 categories |
| 3 | Quantity | Integer | Raw | Min 1; some large orders (up to 87) |
| 4 | Unit Price | Decimal | Raw | USD |
| 5 | Amount | Decimal | Raw | Equals Quantity × Unit Price on every row |
| 6 | Order Date | Date | Raw | |
| 7 | Ship Date | Date | Raw | |
| 8 | Customer Gender | Text | Raw | F, M, O; **135 blanks** |
| 9 | Order Mode | Text | Raw | 5 values |
| 10 | Rating C | Integer | Raw | 1–5 |
| 11 | State | Text | Raw | California |
| 12 | County | Text | Raw | Used for the map |
| 13 | Days to Deliver | Integer | **Calculated** | Ship Date − Order Date |
| 14 | Weeknum | Integer | **Calculated** | `WEEKNUM(Order Date)` |
| 15 | Gender Value | Text | **Calculated** | Female / Male / Other / Unknown |

---

## 3. Cleaning Process

| Step | Check / action | Result |
|---|---|---|
| 1 | Converted the range into an Excel Table named `orders` so formulas and pivots expand automatically | Done |
| 2 | Checked for duplicate rows and duplicate `TX ID` | 0 duplicates |
| 3 | Checked data types (dates as dates, numbers as numbers) | Order Date and Ship Date stored as real dates |
| 4 | Verified `Amount = Quantity × Unit Price` | No differences found |
| 5 | Found blank `Customer Gender` (135 rows) | Kept the rows; labelled as **Unknown** in `Gender Value` instead of deleting orders |
| 6 | Replaced codes F/M/O with readable labels in `Gender Value` (Female, Male, Other, Unknown) so slicer buttons are clear | Done |
| 7 | Created `Days to Deliver` = Ship Date − Order Date (formatted as a number) | Range 0–14 days |
| 8 | Created `Weeknum` = `WEEKNUM(Order Date)` for the 13-week trend | Weeks 1–13 |
| 9 | Checked for missing values in all other columns | None |
| 10 | Saved the cleaned data as `data/ecommerce_cleaned_dataset.xlsx` | Done |

**Formulas used**
```excel
Days to Deliver : =orders[[#This Row],[Ship Date]]-orders[[#This Row],[Order Date]]
Weeknum         : =WEEKNUM(orders[[#This Row],[Order Date]])
```

---

## 4. Analysis Approach

### 4.1 KPIs

| KPI | Aggregation |
|---|---|
| Total Orders | Count of TX ID |
| Total Quantity | Sum of Quantity |
| Total Amount | Sum of Amount |
| Avg. Rating | Average of Rating C |
| Avg. Days to Deliver | Average of Days to Deliver |

### 4.2 Questions → Charts

| Question | Chart | Pivot (sheet `Sheet1`) |
|---|---|---|
| Trend in the last 13 weeks | Line chart, Quantity and Amount on two axes | `pvttrend` |
| How our customers like to buy | Doughnut by Order Mode | `pvtproduct` |
| How many they buy? | Column chart of quantity buckets (1–10, "More than 10") | `PivotTable5` |
| Which products are popular? | Stacked bar by product split by gender | `pvt matrix` |
| Overall gender split | Doughnut | `pvt matrix` |
| Where do our customers live? | Filled map of California counties | `PivotTable7` |
| How long we take to ship | Column chart of Days to Deliver | `PivotTable13` |
| How happy are our customers? | Clustered column, ratings by month | `Pivotratinjg` |

### 4.3 Slicers
- **Order Mode** and **Gender Value** are placed in the left panel of the dashboard.
- Both are connected to every PivotTable through *Report Connections*, so one click updates all KPIs and charts.
- Step-by-step instructions are in the [README](README.md#-how-to-add-slicers-step-by-step).

---

## 5. Final Result

The `Dashboard` sheet contains:
- **Left panel:** 5 KPI cards (Orders, Quantity, Amount, Avg Rating, Days to Deliver) and the 2 slicers.
- **Top row:** Last 13 week Trends, How they like to buy, How many they buy.
- **Middle:** Which products are popular (by gender) and the California county map.
- **Right column:** How long we take to ship, How satisfied are our customers.

Snapshot: `screenshots/dashboard-snapshot.png`

---

## 6. Outcomes and Insights

Computed from the `orders` table:

| Area | Finding |
|---|---|
| Orders | 2,400 orders in 13 weeks |
| Volume and revenue | 11,997 units and $649,019.80 in total amount |
| Satisfaction | Average rating 3.96 (≈ 4.0); 4★ = 1,036, 5★ = 721, 3★ = 495, 2★ = 127, 1★ = 21 |
| Delivery | Average 2.34 days; most common is 1 day (888 orders); maximum 14 days |
| Channels | App 868 · Website 569 · Target.com 453 · Partner App 262 · Instagram 248 |
| Products (orders) | T-Shirts 325 · Jeans 286 · Sneakers 247 · Tank Tops 187 · Bikinis 162 … Jewelry 17 (lowest) |
| Order size | 2 units is most common (540), then 1 (397) and 3 (375) |
| Gender | Female 1,231 · Male 959 · Unknown 135 · Other 75 |
| Location | Los Angeles County 781 orders, then San Diego 330, Sacramento 162, Santa Clara 135, Alameda 132 |
| Weekly trend | Orders rise from 101 (week 1) to 240 (week 6), then hold at roughly 185–240 per week |

**Business takeaways**
1. The mobile **App** is the strongest channel, so app experience and promotions deserve investment.
2. A handful of products (T-Shirts, Jeans, Sneakers) drive volume; low sellers such as Jewelry and Tote Bags could be bundled or promoted.
3. Most customers are satisfied, but about 6% of orders are rated 1–2★; checking these against long delivery times would be a useful follow-up.
4. Demand is concentrated in Los Angeles and San Diego; delivery capacity there matters most.
5. Missing gender (5.6% of orders) should be captured at checkout to improve segmentation.

---

## 7. Validation Notes

- Orders (2,400), average rating (≈ 4.0) and the data-driven figures above match the dataset.
- The dashboard snapshot shows **Quantity 447** and **Amount $31.4k**, while the full `orders` table sums to **11,997 units** and **$649,019.80**; Average Days to Deliver is shown as 2.1 versus 2.34 calculated from the table. These differences usually come from a filter or slicer selection, or a pivot cache that was not refreshed, when the screenshot was taken. Before publishing, click **Data → Refresh All**, clear both slicers, and retake the screenshot so the KPI cards match the table.
- Always keep the slicers cleared when capturing the final snapshot.

---

## 8. Limitations and Next Steps

- Data covers only one state and 13 weeks, so seasonality cannot be assessed.
- No cost or profit field, so profitability cannot be measured.
- Possible extensions: add Product-category and Month slicers, a timeline slicer on Order Date, a profit KPI, and a Power BI version of the same report.

---

## Appendix – Repository Checklist

- [x] Organized file (raw dataset)
- [x] Cleaned file
- [x] Snapshot / screenshot of the report
- [x] Documentation (this file)
- [x] README file
