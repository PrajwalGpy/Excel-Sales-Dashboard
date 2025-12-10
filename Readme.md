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

_(Dashboard 1: Sales Performance)_
![Dashboard Demo](https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20114646.png)

_(Dashboard 2: Customer & Profitability)_
![Dashboard Demo](https://github.com/PrajwalGpy/Excel-Sales-Dashboard/blob/main/assets/Screenshot%202025-12-10%20114823.png)

---

## Tools Used

**Microsoft Excel**: Primary tool for analysis and visualization.

**Pivot Tables**: Used extensively to summarize data for both dashboards.

**VBA & Macros**: To create navigation buttons between dashboards and reset filters.

**Advanced Formulas**: `XLOOKUP`, `IFS`, `TEXT`, and `SUMPRODUCT` for data enrichment.

---

## Steps in Project

✔️ **Data Collection**: Importing raw sales and target data.
✔️ **Data Cleaning**: Handling missing values and standardizing date formats.
✔️ **Data Modeling**: Creating relationships between Transaction, Product, and Customer tables.
✔️ **Dashboard 1 Creation**: Building the Sales Performance view.
✔️ **Dashboard 2 Creation**: Building the Customer & Profitability view.
✔️ **Interactivity**: Adding slicers, timelines, and navigation buttons.
✔️ **Formatting**: Applying a consistent color theme and layout.

---

## Business Requirement

To conduct a **comprehensive analysis of sales data** split into two key areas:

- **Dashboard 1 (Sales Focus)**: Monitor overall revenue, trends, and top-performing products.
- **Dashboard 2 (Customer Focus)**: Understand who the customers are (Age, Gender) and which segments are most profitable.

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

### Dashboard 1: Sales Performance

**1. Monthly Sales Trend (Line Chart)**

- **Objective:** Visualize revenue and profit trends over time (Jan-Dec).
- **Insight:** Identifies seasonal peaks and slow months.
- **Chart Type:** Line Chart with Markers.
<div style="display: flex; justify-content: center; align-items: center; gap: 20px;">
      <img src="https://github.com/PrajwalGpy/Project-PowerBI-AmazonSales-ReviewsAnalysis/blob/main/images/Screenshot%202025-11-16%20115254.png" alt="YTD Sales by Month" width="250" height="200" />
  </div>

**2. Top 5 Products by Revenue (Bar Chart)**

- **Objective:** Highlight the best-selling products.
- **Insight:** Helps in inventory planning for high-demand items.
- **Chart Type:** Clustered Bar Chart.
<div style="display: flex; justify-content: center; align-items: center; gap: 20px;">
      <img src="https://github.com/PrajwalGpy/Project-PowerBI-AmazonSales-ReviewsAnalysis/blob/main/images/Screenshot%202025-11-16%20115254.png" alt="YTD Sales by Month" width="250" height="200" />
  </div>

**3. Sales by Region (Map/Column Chart)**

- **Objective:** Show sales distribution across different geographical locations.
- **Insight:** Pinpoints high-performing regions vs. those needing marketing support.
- **Chart Type:** Map or Column Chart.
<div style="display: flex; justify-content: center; align-items: center; gap: 20px;">
      <img src="https://github.com/PrajwalGpy/Project-PowerBI-AmazonSales-ReviewsAnalysis/blob/main/images/Screenshot%202025-11-16%20115254.png" alt="YTD Sales by Month" width="250" height="200" />
  </div>

**4. Sales by Manager (Bar Chart)**

- **Objective:** Compare performance across different sales representatives.
- **Insight:** Useful for performance appraisals and incentive planning.
- **Chart Type:** Bar Chart.
<div style="display: flex; justify-content: center; align-items: center; gap: 20px;">
      <img src="https://github.com/PrajwalGpy/Project-PowerBI-AmazonSales-ReviewsAnalysis/blob/main/images/Screenshot%202025-11-16%20115254.png" alt="YTD Sales by Month" width="250" height="200" />
  </div>

---

### Dashboard 2: Customer & Profitability

**1. Revenue by Age Group (Column Chart)**

- **Objective:** Segment sales based on customer age brackets (e.g., Youth, Adult, Senior).
- **Insight:** Determines the target audience age range for marketing campaigns.
- **Chart Type:** Column Chart.
<div style="display: flex; justify-content: center; align-items: center; gap: 20px;">
      <img src="https://github.com/PrajwalGpy/Project-PowerBI-AmazonSales-ReviewsAnalysis/blob/main/images/Screenshot%202025-11-16%20115254.png" alt="YTD Sales by Month" width="250" height="200" />
  </div>

**2. Profit by Gender (Donut Chart)**

- **Objective:** Analyze the contribution of Male vs. Female customers to the bottom line.
- **Insight:** Helps tailor gender-specific product messaging.
- **Chart Type:** Donut Chart.
<div style="display: flex; justify-content: center; align-items: center; gap: 20px;">
      <img src="https://github.com/PrajwalGpy/Project-PowerBI-AmazonSales-ReviewsAnalysis/blob/main/images/Screenshot%202025-11-16%20115254.png" alt="YTD Sales by Month" width="250" height="200" />
  </div>

**3. Revenue by Payment Method (Bar Chart)**

- **Objective:** Track preference for Credit Card, Cash, or Online Payment.
- **Insight:** Optimization of payment gateways and understanding customer liquidity.
- **Chart Type:** Bar Chart.
<div style="display: flex; justify-content: center; align-items: center; gap: 20px;">
      <img src="https://github.com/PrajwalGpy/Project-PowerBI-AmazonSales-ReviewsAnalysis/blob/main/images/Screenshot%202025-11-16%20115254.png" alt="YTD Sales by Month" width="250" height="200" />
  </div>

**4. Weekday vs. Weekend Sales (Pie/Donut Chart)**

- **Objective:** Compare sales volume on weekdays versus weekends.
- **Insight:** Informs staffing schedules and weekend promotional strategies.
- **Chart Type:** Pie Chart.
<div style="display: flex; justify-content: center; align-items: center; gap: 20px;">
      <img src="https://github.com/PrajwalGpy/Project-PowerBI-AmazonSales-ReviewsAnalysis/blob/main/images/Screenshot%202025-11-16%20115254.png" alt="YTD Sales by Month" width="250" height="200" />
  </div>

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

1.  Download the file: **`Excel_Sales_Customer_Dashboard.xlsm`**.
2.  Open in **Microsoft Excel** (Enable Macros).
3.  Use the **Navigation Buttons** at the top to switch between "Sales Dashboard" and "Customer Dashboard".
4.  Use Slicers on the left to filter both dashboards simultaneously.

---

## File Details

- **File Name**: `Excel_Sales_Customer_Dashboard.xlsm`
  **Description**: The complete multi-dashboard Excel file.
  [Download File Here](https://www.google.com/search?q=https://github.com/PrajwalGpy/Project-Excel-Sales-Analysis/blob/main/Excel_Sales_Customer_Dashboard.xlsm)

- **File Name**: `Sales_Data.xlsx`
  **Description**: Raw dataset used for the project.
  [Download File Here](https://www.google.com/search?q=https://github.com/PrajwalGpy/Project-Excel-Sales-Analysis/blob/main/Sales_Data.xlsx)

---

## Contact

For any queries or feedback, feel free to reach out:

**Prajwal Gopal Poojary**
Email: `prajwalgpa@gmail.com`
Portfolio: [https://prajwalgopalpoojary.netlify.app](https://prajwalgopalpoojary.netlify.app)
LinkedIn: [https://linkedin.com/in/prajwalgopalpoojary/](https://linkedin.com/in/prajwalgopalpoojary/)

---

## Acknowledgments

Special thanks to **Data with Decision** for the tutorial series.

---
