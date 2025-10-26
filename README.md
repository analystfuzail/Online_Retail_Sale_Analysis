# Online_Retail_Sale_Analysis

🧠 Advanced Excel Practice — Based on Online Retail Dataset
## 🔹 1. Advanced Formulas & Logical Functions

Write a formula to calculate total sales (Quantity * UnitPrice) and dynamically handle blank or negative quantities.

Identify the top 10 most profitable products using formulas only (without PivotTables).

Using SUMIFS, compute total sales per country and month.

Use INDEX + MATCH (or XLOOKUP) to retrieve the UnitPrice of a given product description.

Using IFERROR, build a lookup that handles missing Customer IDs gracefully.

Build a formula to find the first and last purchase date for each customer.

Write a single-cell array formula (or dynamic array formula) to calculate unique customers per country.

Using TEXT, LEFT, MID, RIGHT, extract the country code or part of product description based on pattern.

Identify invoices with negative or duplicate entries using conditional formula logic (e.g., COUNTIFS, ABS, etc.).

Build a formula that returns “High”, “Medium”, or “Low” value customers based on their total spend percentiles (using PERCENTILE.INC).

## 🔹 2. Pivot Tables & Data Analysis

Create a pivot table to summarize sales by country, by month, and by top 5 customers.

Build a calculated field in PivotTable for profit (e.g., assume a fixed cost percentage).

Use Pivot filtering to show only transactions above the average sales.

Show average unit price per product category using Pivot Table grouping.

Create a PivotChart to show monthly sales trends per region.

Combine multiple pivot tables using a Pivot Chart dashboard.

Group PivotTable dates by Month, Quarter, and Year simultaneously.

Use a Slicer to filter by Country or Customer.

Build a PivotTable that dynamically updates when new data is added (using Excel Tables).

Use GETPIVOTDATA to extract values from a PivotTable into a dashboard.

## 🔹 3. Data Modeling & Power Query

Use Power Query to remove duplicates and load only valid transactions (positive quantity and non-null customer IDs).

Split the InvoiceDate into separate Date and Time columns using Power Query.

Merge the Online Retail dataset with another table (e.g., a Country Region table) to enrich analysis.

Create a calendar table in Power Query and establish a relationship with the fact table.

Transform text columns to proper case and trim spaces automatically.

Create a calculated column for Month-Year in Power Query (e.g., “Jan-2025”).

Build a data model with relationships (Customer → Invoice → Product).

Write DAX measures for Total Sales, Average Order Value, and Distinct Customers.

Load only rows with UnitPrice > 0 and Quantity > 0.

Automatically refresh the data model using Power Query parameters.

## 🔹 4. Data Validation & Error Handling

Create a data validation rule that restricts Quantity to positive integers only.

Add a dropdown to select Country dynamically (using named range or table).

Build a dependent dropdown — e.g., selecting a Country filters Customer IDs available.

Highlight invalid records using Conditional Formatting + Data Validation (e.g., negative UnitPrice).

Set up a validation rule that ensures InvoiceNo starts with a letter “C” for credit notes.

Create a warning for any blank description or duplicate invoice entry.

Use a custom validation formula to ensure UnitPrice × Quantity < $10,000.

Protect and lock only the data input range, leaving analysis cells editable.

Build a form with validated dropdowns for manual transaction entry.

Use a macro-enabled button to clear validated fields after each new entry.

## 🔹 5. Dashboarding & Visualization

Build a dynamic sales dashboard showing KPIs — Total Sales, No. of Customers, Avg Order Value, etc.

Create a Top 5 products by revenue chart that updates when Country is changed.

Add a timeline slicer for monthly sales tracking.

Use sparklines to show trends in customer spending.

Apply conditional formatting heatmaps for sales by region and product.

Use camera tool or linked pictures for dashboard interactivity.

Combine scroll bars or dropdowns to dynamically switch between KPIs (sales, quantity, etc.).

Build a what-if analysis scenario (e.g., UnitPrice increase by 10%).

Create a dashboard summary using Form Controls or ActiveX Controls.

Design a fully automated report refresh (using Power Query + PivotTable refresh + macros).

## 🔹 6. Automation & VBA Integration (Optional Bonus)

Write a VBA macro to refresh all PivotTables and Power Query connections.

Create a macro to export daily sales summary into a new Excel file.

Use VBA to highlight invoices exceeding $10,000.

Automate sending an email summary (e.g., top 5 countries’ revenue).

Build a user form for manual transaction entry with validation rules.
