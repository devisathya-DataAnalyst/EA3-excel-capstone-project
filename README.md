# EA3-excel-capstone-project
Data exploration and Analysis using Excel
# E-commerce Sales Data Analysis & Visualization

## Project Overview

This project focuses on **Data Cleaning, Analysis & Visualization using Microsoft Excel**. The objective is to transform raw e-commerce sales data into meaningful business insights through data cleaning, Excel formulas, statistical analysis, data summarization, and interactive dashboards.

### Dataset Overview

| Metric | Details |
|---|---:|
| Sales Transactions | 2,000 |
| Data Period | 2023–2025 |
| Customers | 500 |
| Products | 100 |
| Stores | 20 |

---

## 1. Data Preparation and Quality Decisions

The workbook preserves the four original source tables and provides separate cleaned tables for analysis.

### Data Cleaning Activities

- Checked for duplicate records and duplicate primary keys. No duplicates were identified.
- Standardized whitespace and text casing across relevant columns.
- Standardized country names by converting `United States of America` to `USA`.
- Corrected the Product ID `P-3` to `PROD3`.
- Used Power Query Editor to replace Customer ID `C-1464` with `CUST1464`.
- Reviewed all columns for relevance. No irrelevant columns were identified for removal.

### Missing Value Handling

| Dataset | Missing Data | Treatment |
|---|---|---|
| Customer_Dim | 4 Loyalty Level values | Replaced with `Unknown` |
| Product_Dim | 6 Cost values | Category median (overall median as fallback) |
| Product_Dim | 3 Stock values | Overall median |
| Sales_Fact | 5 Quantity values | Median, rounded to whole units |
| Sales_Fact | 3 Unit Price values | Column median |
| Sales_Fact | 4 Discount values | Column median |
| Sales_Fact | Total Amount | Recalculated for all records |

**Total Amount Calculation:**

```excel
Total_Amount = Quantity * Unit_Price * (1 - Discount)
```

These cleaning and transformation steps retain the original transaction records while addressing missing and inconsistent values. Imputed values should be treated as estimates rather than observed values.

Screenshots documenting the cleaning and transformation activities are included with the submission.

---

## 2. Key Sales Findings

### Business Performance KPIs

| KPI | Result |
|---|---:|
| **Net Revenue** | **$2,173,211.54** |
| Total Transactions | 2,000 |
| Units Sold | 4,958 |
| Average Order Value | $1,086.61 |
| Highest Revenue Category | Sports |
| Sports Revenue | $542,010.12 |
| Highest Revenue Region | North |
| North Region Revenue | $643,578.63 |

### Revenue by Product Category

| Product Category | Revenue |
|---|---:|
| Sports | $542,010.12 |
| Electronics | $505,378.31 |
| Home | $424,794.97 |
| Clothing | $378,379.49 |
| Beauty | $300,900.06 |

### Key Business Insights

- **Sports** generated the highest revenue, followed by Electronics.
- **North** was the strongest-performing store region, generating $643,578.63.
- South and East regions recorded relatively close revenue performance.
- West contributed the lowest revenue among the four regions.

These findings represent descriptive analysis of the supplied transaction data and do not establish the underlying reasons for differences in performance.

---

## 3. Statistical Profile and Analysis

Descriptive statistics were calculated using Excel statistical functions.

| Statistical Measure | Value |
|---|---:|
| Mean | $1,086.61 |
| Median | $868.72 |
| Mode | $355.24 |
| Standard Deviation (Sample) | $818.48 |
| Minimum | $18.46 |
| Maximum | $3,908.24 |

### Statistical Interpretation

The mean transaction value ($1,086.61) is higher than the median ($868.72). This indicates that larger transactions pull the average order value upward.

The standard deviation of $818.48 indicates considerable variation in transaction values.

---

## 4. Excel Formulas and Functions

The project demonstrates the application of Excel formulas for calculations, data transformation, and analysis.

### Functions Used

| Function | Application |
|---|---|
| `SUM()` | Calculate total revenue and units sold |
| `IF()` | Apply conditional logic |
| `COUNTIF()` | Count transactions matching specific criteria |
| `AVERAGE()` | Calculate average transaction value |
| `TEXT()` | Format dates and numerical values |
| `XLOOKUP()` | Retrieve customer and product information |
| `VLOOKUP()` | Retrieve store information |
| `MEDIAN()` | Handle missing values and calculate median |
| `MODE.SNGL()` | Identify the most frequent transaction value |
| `STDEV.S()` | Calculate sample standard deviation |

### Combining Tables

The XLOOKUP and VLOOKUP functions were used to enrich sales transaction data by retrieving relevant information from customer, product, and store dimension tables.

---

## 5. Data Summarization

Pivot-style summaries were created to analyze sales performance across multiple dimensions.

### Analysis Dimensions

- Revenue by Product Category
- Revenue by Store Region
- Revenue by Payment Type
- Monthly Revenue Trends

The `Pivot_Summaries` worksheet contains the summarized results used for analysis and visualization.

---

## 6. Data Visualization and Interactive Dashboard

An interactive Excel dashboard was developed to present key business performance indicators and sales trends.

### Visualization Components

- **Column Chart:** Revenue by Product Category.
- **Pie Chart:** Revenue distribution by Payment Type.
- **Line Chart:** Monthly Revenue Trends.
- **Dashboard Filters:** Region and Category dropdown filters.
- **KPI Cards:** Revenue, Transactions, Average Order Value, and Units Sold.

The dashboard allows users to explore sales performance by selecting different regions and product categories.

---

## 7. Project Deliverables

The project includes the following deliverables:

| Deliverable | Description |
|---|---|
| Excel Workbook | Raw data, cleaned data, formulas, statistics, summaries, and dashboard |
| Cleaning Log | Documented data quality issues and corrective actions |
| Screenshots | Evidence of data cleaning and transformation steps |
| Findings Summary | Business insights and statistical analysis |

---

## 8. Conclusion

This project demonstrates the end-to-end application of Microsoft Excel for e-commerce data analysis.

Through data cleaning, transformation, statistical analysis, lookup functions, data summarization, and visualization, raw transactional data was converted into meaningful business insights.

The analysis identified Sports as the highest revenue-generating category and North as the strongest-performing region. The interactive dashboard provides a consolidated view of business performance and supports data-driven exploration.

---

## Tools and Technologies

- Microsoft Excel
- Power Query Editor
- Excel Tables
- Excel Formulas and Functions
- XLOOKUP / VLOOKUP
- Pivot-style Summaries
- Excel Charts and Interactive Dashboards

---

**Project:** E-commerce Sales Data Analysis & Visualization  
**Primary Tool:** Microsoft Excel  
**Focus Areas:** Data Cleaning | Data Analysis | Statistics | Data Visualization | Business Insights
