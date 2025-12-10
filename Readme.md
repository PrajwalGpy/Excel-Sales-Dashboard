# Excel Sales & Customer Insights Dashboard

This project uses **Microsoft Excel** to create a dual-dashboard solution: a **Sales Performance Dashboard** for tracking high-level metrics and a **Customer & Profitability Dashboard** for deep-dive demographics and margin analysis.
The goal is to provide a complete view of the business, from revenue trends to customer behavior, using dynamic and interactive visualizations.

---

## Dashboard Demo

![Dashboard Demo](https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Recording%202025-12-10%20115226.gif)
![Dashboard Demo](https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Recording%202025-12-10%20115338.gif)

---

## Key Features

📌 **Dual-Dashboard Layout**: Separate views for **Sales Overview** and **Customer/Profit Analysis**.

📌 **Dynamic Metric Switching**: Toggle visuals between **Revenue**, **Profit**, and **Quantity** using VBA macros.

📌 **Customer Demographics**: Deep analysis of sales by **Age Group** and **Gender**.

📌 **Advanced Interactivity**: Custom slicers, timelines, and "Clear Filter" macros for a seamless user experience.

---

_(Dashboard 1: Customer & Profitability)_
![Dashboard Demo](https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20114646.png)

_(Dashboard 2: Sales Performance)_
![Dashboard Demo](https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20114823.png)

---

## Tools Used

**Microsoft Excel**: Primary tool for analysis and visualization.

**Pivot Tables**: Used extensively to summarize data for both dashboards.

**VBA & Macros**: To create navigation buttons between dashboards and reset filters.

**Advanced Formulas**: `XLOOKUP`, `IFS`, `TEXT`, and `SUMPRODUCT` for data enrichment.

---

## Steps in Project

✔️ Data Collection: Importing raw sales and target data.
✔️ Data Cleaning: Handling missing values and standardizing date formats.
✔️ Data Modeling: Creating relationships between Transaction, Product, and Customer tables.
✔️ Dashboard 1 Creation: Building the Sales Performance view.
✔️ Dashboard 2 Creation: Building the Customer & Profitability view.
✔️ Interactivity: Adding slicers, timelines, and navigation buttons.
✔️ Formatting: Applying a consistent color theme and layout.

---

## Business Requirement

To conduct a **comprehensive analysis of sales data** split into two key areas:

- **Dashboard 2 (Customer Focus)**: Understand who the customers are (Age, Gender) and which segments are most profitable.
- **Dashboard 1 (Sales Focus)**: Monitor overall revenue, trends, and top-performing products.

---

## KPI’s Requirements

**1. Total Revenue:**
Total income generated from sales transactions.

**2. Total Profit:**
Revenue minus the Cost of Goods Sold (COGS).

**3. Total Transactions:**
The count of unique orders processed.

**4. Profit Margin %:**
The ratio of profit to revenue, indicating business efficiency.

![KPIs Screenshot](https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20114703.png)

---

## Chart’s Requirements

### Dashboard 1: Customer & Profitability

<ol>

<h3><li> Revenue by Age Group (Column Chart):</li></h3> <ul> <li><b><ins>Objective:</ins></b> Segment sales based on customer age brackets (e.g., Youth, Adult, Senior).</li> <li><b>Insight:</b> Determines the target audience age range for marketing campaigns.</li> <li><b>Chart Type:</b> Column Chart.</li>

<div style="display: flex; justify-content: center; align-items: center; gap: 20px;"> <img src="https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20115033.png" alt="Revenue by Age Group" width="250" height="200" /> </div> </ul>

<h3><li> Profit by Gender (Donut Chart):</li></h3> <ul> <li><b><ins>Objective:</ins></b> Analyze the contribution of Male vs. Female customers to the bottom line.</li> <li><b>Insight:</b> Helps tailor gender-specific product messaging.</li> <li><b>Chart Type:</b> Donut Chart.</li>

<div style="display: flex; justify-content: center; align-items: center; gap: 20px;"> <img src="https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20115113.png" alt="Profit by Gender" width="250" height="200" /> </div> </ul>

<h3><li> Revenue by Payment Method (Bar Chart):</li></h3> <ul> <li><b><ins>Objective:</ins></b> Track preference for Credit Card, Cash, or Online Payment.</li> <li><b>Insight:</b> Optimization of payment gateways and understanding customer liquidity.</li> <li><b>Chart Type:</b> Bar Chart.</li>

<div style="display: flex; justify-content: center; align-items: center; gap: 20px;"> <img src="https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20115104.png" alt="Revenue by Payment Method" width="250" height="200" /> </div> </ul>

<h3><li> Sales by Region (Map/Column Chart):</li></h3> <ul> <li><b><ins>Objective:</ins></b> Show sales distribution across different geographical locations.</li> <li><b>Insight:</b> Pinpoints high-performing regions vs. those needing marketing support.</li> <li><b>Chart Type:</b> Map or Column Chart.</li>

<div style="display: flex; justify-content: center; align-items: center; gap: 20px;"> <img src="https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20115055.png" alt="Sales by Region" width="250" height="200" /> </div> </ul>

## </ol>

### Dashboard 2: Sales Performance

<ol>

<h3><li> Revenue Trend by Quarter (Column Chart):</li></h3> <ul> <li><b><ins>Objective:</ins></b> Visualize the distribution of revenue across the four quarters of the year.</li> <li><b>Insight:</b> Q2 is the strongest performing quarter (35.0%), indicating a mid-year peak, while Q4 shows a significant drop (10.3%).</li> <li><b>Chart Type:</b> Column Chart.</li>

<div style="display: flex; justify-content: center; align-items: center; gap: 20px;"> <img src="https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20114950.png" alt="Revenue Trend by Quarter" width="250" height="200" /> </div> </ul>

<h3><li> Monthly Revenue Fluctuation (Line Chart):</li></h3> <ul> <li><b><ins>Objective:</ins></b> Track the revenue trajectory over the entire year to identify seasonality and sharp declines.</li> <li><b>Insight:</b> Jan, Mar, and Jul are top-performing months, but there is a drastic drop in revenue starting in August, continuing through December.</li> <li><b>Chart Type:</b> Line Chart with Data Markers.</li>

<div style="display: flex; justify-content: center; align-items: center; gap: 20px;"> <img src="https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20114942.png" width="250" height="200" /> </div> </ul>

<h3><li> Daily Sales Trends (Column Chart):</li></h3> <ul> <li><b><ins>Objective:</ins></b> Analyze sales performance by day of the week to understand daily customer purchasing habits.</li> <li><b>Insight:</b> Sales are relatively consistent throughout the week, with marginal peaks on Tuesdays and Thursdays.</li> <li><b>Chart Type:</b> Clustered Column Chart.</li>

<div style="display: flex; justify-content: center; align-items: center; gap: 20px;"> <img src="https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20115012.png" alt="Daily Sales Trends" width="250" height="200" /> </div> </ul>

<h3><li> Weekend vs. Weekday Distribution (Donut Charts):</li></h3> <ul> <li><b><ins>Objective:</ins></b> Compare the volume of business conducted during the work week versus the weekend.</li> <li><b>Insight:</b> The majority of sales (72.6%) occur during weekdays, suggesting B2B activity or weekday-specific shopping patterns.</li> <li><b>Chart Type:</b> Radial / Donut Chart.</li>

<div style="display: flex; justify-content: center; align-items: center; gap: 20px;"> <img src="https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20115001.png" width="250" height="200" /> </div> </ul>

</ol>

---

## Dashboard Insights

### Highlights:

1.  **Dual Perspective**: Separating Sales and Customer data prevents information overload and allows for targeted analysis.
2.  **Demographic Targeting**: Dashboard 2 clearly shows that the **Adult (30-50)** age group drives 60% of revenue.
3.  **Seasonality**: Dashboard 1 reveals a consistent sales dip in **Q3**, suggesting a need for mid-year promotions.

### Key Insights:

1.  **Sales Trends:**
    Monthly trends indicate strong end-of-year performance, likely driven by holiday sales.

2.  **Customer Behavior:**
    Female customers tend to buy higher-margin items, while Male customers purchase higher quantities of lower-margin items.

3.  **Payment Preferences:**
    Digital payments account for the majority of high-value transactions.

---

## 📂 How to Use

1.  Download the file: **`Sales Performance Dataset_macro.xlsm`**.
2.  Open in **Microsoft Excel** (Enable Macros).
3.  Use the **Navigation Buttons** at the top to switch between "Sales Dashboard" and "Customer Dashboard".
4.  Use Slicers on the left to filter both dashboards simultaneously.

---

## File Details

- **File Name**: `Sales Performance Dataset_macro.xlsm`
  **Description**: The complete multi-dashboard Excel file.
  [Download File Here](https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/Sales%20Performance%20Dataset_macro.xlsm)

- **File Name**: `SalesPractice Dataset.xlsx`
  **Description**: Raw dataset used for the project.
  [Download File Here](https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/SalesPractice%20Dataset.xlsx)

---

##  Contact

For any queries or feedback, feel free to reach out:

**Prajwal Gopal Poojary**  
 Email: `prajwalgpa@gmail.com`  
 Portfolio: <https://prajwalgo](https://prajwalgp.me/>  
 LinkedIn: <https://linkedin.com/in/prajwalgopalpoojary/>

---

## Acknowledgments

Special thanks to **Data with Decision** for the tutorial series.

---




