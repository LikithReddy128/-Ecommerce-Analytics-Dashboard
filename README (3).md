# 🛒 Ecommerce Excel Dashboard

> An interactive Excel dashboard that turns 2,400 raw ecommerce orders (California, Q1 2025) into answers about sales trends, customer buying behaviour, product popularity, delivery speed and customer satisfaction.

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=white)
![Pivot Tables](https://img.shields.io/badge/Technique-PivotTables%20%26%20Slicers-blue)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen)

---

## 📌 Table of Contents
1. [Project Overview](#-project-overview)
2. [Repository Structure](#-repository-structure)
3. [Dataset at a Glance](#-dataset-at-a-glance)
4. [KPIs](#-kpis)
5. [Business Questions & Charts](#-business-questions--charts)
6. [Slicers](#-slicers)
7. [How to Build This Dashboard (Step by Step)](#-how-to-build-this-dashboard-step-by-step)
8. [How to Add Slicers (Step by Step)](#-how-to-add-slicers-step-by-step)
9. [How to Create This GitHub Repository (Step by Step)](#-how-to-create-this-github-repository-step-by-step)
10. [How to Use the Dashboard](#-how-to-use-the-dashboard)
11. [Key Insights](#-key-insights)
12. [Tools & Skills](#-tools--skills)
13. [Documentation](#-documentation)

---

## 📖 Project Overview

| | |
|---|---|
| **Project name** | Ecommerce Excel Dashboard |
| **Goal** | Give a business owner one screen that answers 8 key questions about orders, customers, products and service quality |
| **Data period** | 1 Jan 2025 – 29 Mar 2025 (13 weeks) |
| **Region** | California, USA (58 counties) |
| **Tool** | Microsoft Excel (Tables, PivotTables, PivotCharts, Slicers, Map chart) |

---

## 📁 Repository Structure

Create exactly this layout in your repository (the folder names follow the checklist in the project notes: organized file, cleaned file, snapshot, documentation, readme).

```
ecommerce-excel-dashboard/
│
├── README.md                          ← you are here (everything in one place)
├── DOCUMENTATION.md                   ← detailed project documentation
│
├── data/
│   ├── ecommerce-blank_dataset.xlsx           ← raw / organized dataset (2,400 rows)
│   └── ecommerce_cleaned_dataset.xlsx         ← cleaned file (optional export of the Data sheet)
│
├── dashboard/
│   └── ecommerce-blank_excel_project.xlsx     ← final workbook: Data, Questions & KPIs, Dashboard, Sheet1 (pivots)
│
└── screenshots/
    └── dashboard-snapshot.png                 ← screenshot of the finished dashboard
```

**Workbook sheets**

| Sheet | Purpose |
|---|---|
| `Data` | Excel Table named **`orders`** (A1:O2401) |
| `Questions & KPIs` | The 5 KPIs and 8 business questions the dashboard must answer |
| `Dashboard` | Final report: KPI cards, 8 charts, 2 slicers |
| `Sheet1` | Back-end sheet holding all 9 PivotTables that feed the charts |

---

## 🗂 Dataset at a Glance

**Rows:** 2,400  |  **Columns:** 15  |  **Table name:** `orders`

| Column | Description |
|---|---|
| TX ID | Unique transaction id (TX500001 …) |
| Product | One of 20 products (T-Shirts, Jeans, Sneakers, …) |
| Quantity | Units in the order |
| Unit Price | Price per unit (USD) |
| Amount | Quantity × Unit Price |
| Order Date | Date the order was placed |
| Ship Date | Date the order shipped |
| Customer Gender | F / M / O, or blank |
| Order Mode | App, Website, Target.com, Partner App, Instagram |
| Rating C | Customer rating 1–5 |
| State | California |
| County | 58 California counties |
| **Days to Deliver** *(calculated)* | `=[@[Ship Date]]-[@[Order Date]]` |
| **Weeknum** *(calculated)* | `=WEEKNUM([@[Order Date]])` |
| **Gender Value** *(calculated)* | Readable label: Female / Male / Other / Unknown (blank → Unknown) |

---

## 🎯 KPIs

The five KPI cards at the top-left of the dashboard (built from the PivotTable `pvtsummary`):

| KPI | Calculation | Pivot value setting |
|---|---|---|
| 🛒 **Orders** | Count of `TX ID` | Count |
| 📦 **Quantity** | Sum of `Quantity` | Sum |
| 💰 **Amount** | Sum of `Amount` | Sum |
| ⭐ **Avg Rating** | Average of `Rating C` | Average |
| 📅 **Days to Deliver** | Average of `Days to Deliver` | Average |

> Screenshot values: Orders **2,400**, Avg Rating **4.0**, Days to Deliver **2.1**. See the note in [DOCUMENTATION.md](DOCUMENTATION.md#7-validation-notes) about the Quantity and Amount cards.

---

## ❓ Business Questions & Charts

| # | Question (from the `Questions & KPIs` sheet) | Dashboard title | Chart type | Pivot used |
|---|---|---|---|---|
| 1 | Trend in the last 13 weeks | **Last 13 week Trends – Qty & Amount** | Line chart (combo, 2 axes) | `pvttrend` |
| 2 | How our customers like to buy | **How they like to buy** | Doughnut | `pvtproduct` / order-mode pivot |
| 3 | How many they buy? | **How many they buy!** | Column chart (quantity buckets 1–10, "More than 10") | `PivotTable5` |
| 4 | Which products are popular? | **Which products are Popular? breakdown by Gender** | Stacked bar by product + gender doughnut | `pvt matrix` |
| 5 | Overall gender split | Doughnut inside the product chart | Doughnut | `pvt matrix` |
| 6 | Where do our customers live? | **Where do our customers live?** | Filled map (California counties) | `PivotTable7` |
| 7 | How long we take to ship the orders | **How long we take to ship** | Column chart (days 0–14) | `PivotTable13` |
| 8 | How happy are our customers? | **How satisfied are our customers?** | Clustered column (rating by month) | `Pivotratinjg` |

---

## 🎛 Slicers

| Slicer | Field | Connected to |
|---|---|---|
| **Order Mode** | `Order Mode` | All PivotTables / charts |
| **Gender Value** | `Gender Value` | All PivotTables / charts |

Clicking any button in a slicer filters **every** KPI card and chart at the same time (Ctrl+click for multi-select).

---

## 🔧 How to Build This Dashboard (Step by Step)

### Step 1 – Prepare the data
1. Open `ecommerce-blank_dataset.xlsx`, click any cell in the data, press **Ctrl + T** → tick *My table has headers* → OK.
2. **Table Design → Table Name** → type `orders`.
3. Check columns: dates formatted as dates, `Amount` as currency, no duplicate `TX ID`.
4. Add the three calculated columns:
   - `Days to Deliver` → `=orders[[#This Row],[Ship Date]]-orders[[#This Row],[Order Date]]` (format as **Number**, not Date)
   - `Weeknum` → `=WEEKNUM(orders[[#This Row],[Order Date]])`
   - `Gender Value` → map `F → Female`, `M → Male`, `O → Other`, blank → `Unknown`
5. Save a copy of the cleaned sheet as `ecommerce_cleaned_dataset.xlsx`.

### Step 2 – Create the sheets
Add three more sheets: **Questions & KPIs**, **Dashboard**, **Sheet1** (back-end). Type the 5 KPIs and 8 questions on *Questions & KPIs* so the dashboard has a clear brief.

### Step 3 – Build the KPI pivot (`pvtsummary`)
1. On `Sheet1`: **Insert → PivotTable** → Table/Range `orders` → Existing worksheet.
2. Drag to **Values**: `TX ID` (Count), `Quantity` (Sum), `Amount` (Sum), `Rating C` (Average), `Days to Deliver` (Average).
3. Rename the pivot: **PivotTable Analyze → PivotTable Name** → `pvtsummary`.

### Step 4 – Build the pivots for each chart

| Pivot name | Rows | Columns | Values |
|---|---|---|---|
| `pvttrend` | `Weeknum` | – | Sum of Quantity, Sum of Amount |
| `pvtproduct` | `Order Mode` | – | Count of TX ID |
| `PivotTable5` | `Quantity` (group 1–10 + "More than 10") | – | Count of TX ID |
| `pvt matrix` | `Product` | `Gender Value` | Sum of Quantity (stacked bar) / Sum of Amount, *Show Values As → % of Grand Total* |
| `PivotTable7` | `County` | – | Sum of Quantity |
| `PivotTable13` | `Days to Deliver` | – | Count of TX ID |
| `Pivotratinjg` | `Order Date` grouped by **Months** | `Rating C` | Count of TX ID |

### Step 5 – Insert the charts
For each pivot: click inside it → **PivotTable Analyze → PivotChart** → choose the chart type from the table in [Business Questions & Charts](#-business-questions--charts).
- **Trend chart:** Line chart; move *Sum of Amount* to **Secondary Axis** (Format Data Series → Secondary Axis).
- **Map:** Select the County pivot → **Insert → Maps → Filled Map**.
- Delete chart titles' default text and type the dashboard titles from the table; remove field buttons you do not need (**PivotChart Analyze → Field Buttons**).

### Step 6 – Design the Dashboard sheet
1. Go to `Dashboard`; **View → untick Gridlines** and **Headings**.
2. Cut each chart from `Sheet1` and paste it on `Dashboard`.
3. Left column: KPI cards (a rounded rectangle + icon + a cell link to the pivot value using `=GETPIVOTDATA(...)`, or simply `=Sheet1!C5`).
4. Use a consistent colour palette (peach, green, blue-grey, brown, orange as in the snapshot) and clear chart titles.
5. Align everything with **Page Layout → Align** and leave even gaps between charts.

---

## 🎚 How to Add Slicers (Step by Step)

**Slicer 1 – Order Mode**
1. On `Sheet1`, click any cell inside a PivotTable (e.g. `pvtsummary`).
2. **PivotTable Analyze → Insert Slicer**.
3. Tick **Order Mode** → **OK**.
4. Cut the slicer and paste it on the `Dashboard` sheet (left panel).

**Slicer 2 – Gender Value**
1. Click inside a PivotTable again → **Insert Slicer**.
2. Tick **Gender Value** → **OK** → move it to the `Dashboard` sheet under the first slicer.

**Connect each slicer to ALL pivots (very important)**
1. Right-click the slicer → **Report Connections** (Excel 2013 and later; named *PivotTable Connections* on some versions).
2. Tick **every** PivotTable: `pvtsummary`, `pvttrend`, `pvtproduct`, `PivotTable5`, `pvt matrix`, `PivotTable7`, `PivotTable13`, `Pivotratinjg`, … → **OK**.
3. Repeat for the second slicer.

**Style and test**
1. Select the slicer → **Slicer tab → Columns** = 1 (vertical buttons) and adjust width/height.
2. Choose a slicer style that matches the dashboard colours.
3. Test: click *App* → all KPIs and charts must change. Click the clear-filter icon (top-right of the slicer) to reset.

> ⚠️ If a chart does not react, that pivot is not ticked in *Report Connections*.

---

## 🚀 How to Create This GitHub Repository (Step by Step)

### Option A – Using the GitHub website (no coding)
1. Sign in at **github.com** → click **+ (top-right) → New repository**.
2. **Repository name:** `ecommerce-excel-dashboard`
3. **Description:** *Interactive Excel dashboard with KPIs, PivotCharts and slicers analysing 2,400 ecommerce orders.*
4. Choose **Public**, tick **Add a README file**, **Create repository**.
5. Click **Add file → Upload files**, drag in the `data/`, `dashboard/` and `screenshots/` folders plus `DOCUMENTATION.md`.
   *(To keep folders on the web uploader, drag the whole folder, or use Option B.)*
6. Replace the auto-created `README.md` by uploading this one (**Add file → Upload files → Commit changes**).
7. Commit message: `Add dashboard workbook, data, screenshot and documentation`.

### Option B – Using Git (professional workflow)
```bash
# 1. One-time setup
git config --global user.name  "Your Name"
git config --global user.email "you@example.com"

# 2. Create the local project folder
mkdir ecommerce-excel-dashboard && cd ecommerce-excel-dashboard
mkdir data dashboard screenshots
# copy your files into those folders, plus README.md and DOCUMENTATION.md

# 3. Initialise and commit
git init
git add .
git commit -m "Initial commit: ecommerce Excel dashboard, data, screenshot and docs"

# 4. Connect to GitHub (create the empty repo on github.com first)
git branch -M main
git remote add origin https://github.com/<your-username>/ecommerce-excel-dashboard.git
git push -u origin main
```

### Take and add the screenshot
1. Open `Dashboard`, press **Win + Shift + S** (Windows) or **Cmd + Shift + 4** (Mac) and capture the full dashboard.
2. Save as `screenshots/dashboard-snapshot.png`.
3. In this README add near the top:
   ```markdown
   ![Dashboard Snapshot](screenshots/dashboard-snapshot.png)
   ```

### Professional finishing touches
- Add **Topics** (⚙️ next to *About*): `excel`, `dashboard`, `data-analysis`, `pivot-table`, `slicers`, `ecommerce`.
- Pin the repository on your profile.
- Use clear commit messages (`Add`, `Fix`, `Update` + what changed).
- Optionally add a `.gitignore` containing `~$*.xlsx` to skip Excel temp files.

---

## 🖱 How to Use the Dashboard
1. Download `dashboard/ecommerce-blank_excel_project.xlsx` and open it in Microsoft Excel (desktop version – the map chart and slicers need Excel 2016 or later).
2. Go to the **Dashboard** sheet.
3. Click buttons in **Order Mode** and **Gender Value** to filter everything. Ctrl+click for multiple selections.
4. Use the clear-filter icon on a slicer to reset.

---

## 💡 Key Insights

Figures below are computed directly from the `orders` table (2,400 orders, 1 Jan–29 Mar 2025).

- **Best-selling products by orders:** T-Shirts (325), Jeans (286) and Sneakers (247) lead; Jewelry (17) is the lowest.
- **Favourite channel:** the **App** is the top order mode (868 orders, 36%), followed by Website (569) and Target.com (453).
- **Small baskets dominate:** the most common order size is **2 units** (540 orders), then 3 (375) and 1 (397); a long tail goes up to 87 units.
- **Customers are mostly female:** Female 1,231 (51%), Male 959 (40%), Other 75, Unknown 135 (blank in the raw data).
- **Concentrated in Southern California:** Los Angeles County alone has 781 orders (33%), then San Diego (330) and Sacramento (162).
- **Fast shipping:** the typical order ships in **1 day** (888 orders); average is about **2.3 days**, with a small tail of slow orders up to 14 days.
- **Customers are satisfied:** average rating ≈ **4.0 / 5**; 4★ (1,036) and 5★ (721) make up about 73% of ratings, while only 148 orders (6%) are rated 1–2★.
- **Growth trend:** weekly order counts rise from 101 in week 1 to 240 in week 6 and stay around 185–240 afterwards.

---

## 🧰 Tools & Skills
Microsoft Excel · Excel Tables · Calculated columns · PivotTables · PivotCharts · Slicers & Report Connections · Combo charts · Filled Map chart · Dashboard design · Git & GitHub

---

## 📚 Documentation
Full write-up (aim, dataset description, cleaning process, final result, outcomes, insights, validation notes): **[DOCUMENTATION.md](DOCUMENTATION.md)**

---

## 👤 Author
**Your Name** – [GitHub](https://github.com/your-username) · [LinkedIn](https://www.linkedin.com/in/your-profile)

## 📄 License
This project is shared for learning and portfolio purposes.
