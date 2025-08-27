# Coffee_Sales_DataAnalysis_MSExcel

## 1) Introduction
This project is about turning raw coffee sales data into a clean, interactive Excel dashboard. I built it end-to-end—starting with data cleaning and lookups, then adding calculated fields, and finally creating PivotTables, charts, and slicers for interactivity.
The final dashboard shows:
- **Sales trends over time** by coffee type (line chart)  
- **Sales by country** (bar chart)  
- **Top 5 customers** (bar chart)  
- **Interactive controls**: timeline (Order Date) and slicers (Roast Type, Size, Loyalty Card)
  
The goal was to practice real analyst workflows—data prep, modeling, and visualization—and to show how advanced Excel can deliver quick, self-serve insights for stakeholders.

## 2) Data Model & Sources  

I structured the data in a simple star-schema style to keep things clean and flexible:  

- **Orders (fact table):**  
  - Order ID, Order Date, Customer ID, Product ID, Quantity  
  - Plus enriched columns (Sales, Coffee Type Name, Roast Type Name, etc.)  

- **Customers (dimension):**  
  - Customer ID (key), Name, Email, Country, Loyalty Card  

- **Products (dimension):**  
  - Product ID (key), Coffee Type (short code), Roast Type (M/L/D), Size (0.2 / 0.5 / 1 / 2.5), Unit Price, Price per 100g, Profit  

**Why this structure?**  
This mirrors a classic **star schema** used in analytics: one fact table for transactions, with separate dimension tables for lookup attributes. It keeps the model tidy, avoids duplication, and makes PivotTables easier to build and maintain.  

## 3) Data Gathering & Preparation 

To populate the Orders table with customer and product details, I used modern lookup formulas in Excel. This allowed me to keep the model structured (fact + dimension tables) and avoid repeating the same information everywhere.  

### 3.1 Customer attributes with XLOOKUP  
I pulled **Name**, **Email**, and **Country** from the Customers table into the Orders sheet using `XLOOKUP`:  
= XLOOKUP([@CustomerID], Customers[CustomerID], Customers[CustomerName], "", 0)  
= XLOOKUP([@CustomerID], Customers[CustomerID], Customers[Email], "", 0)  
= XLOOKUP([@CustomerID], Customers[CustomerID], Customers[Country], "", 0) ``` 
For Email, I wrapped the formula in an IF to avoid showing 0 when no email was available:

= IF(XLOOKUP([@CustomerID], Customers[CustomerID], Customers[Email], "", 0)=0, "", 
     XLOOKUP([@CustomerID], Customers[CustomerID], Customers[Email], "", 0))

Why XLOOKUP?
It uses exact match by default, which prevents mistakes.
The syntax is cleaner than INDEX+MATCH when pulling one column at a time.
It has an optional “if not found” argument, which makes the formula more robust.


### 3.2 Product attributes with a dynamic INDEX/MATCH

For product details like Coffee Type, Roast Type, Size, and Unit Price, I used a single INDEX/MATCH formula that could be filled across and down:
= INDEX(Products!$A$1:$H$999,
        MATCH($D2, Products!$A$1:$A$999, 0),
        MATCH(I$1, Products!$A$1:$H$1, 0)) 

$D2 → Product ID in the current row (row relative, column locked).
I$1 → Column header above (row locked, column relative), so dragging right automatically switches to the correct attribute.

Why INDEX/MATCH here?
One formula covers all product attributes by using column headers as references.
It’s more flexible than VLOOKUP because it doesn’t break if columns are inserted or moved—the formula finds the column by name instead of relying on a fixed index.

## 4) New Calculated Columns (How & Why)  

### 4.1 Sales  
=[@UnitPrice] * [@Quantity]
Why: A simple, explicit measure for Pivots and charts. Keeps the “business math” in the model instead of embedding it in visuals.

### 4.2 Renaming (for better readibility)
Coffee Type Name from short codes (ROB, EXE, ARA, LIB) using nested IF:

=IF([@[CoffeeType]]="ROB","Robusta",
 IF([@[CoffeeType]]="EXE","Excelsa",
 IF([@[CoffeeType]]="ARA","Arabica",
 IF([@[CoffeeType]]="LIB","Liberica",""))))

Roast Type Name from M/L/D → Medium/Light/Dark:
=IF([@[RoastType]]="M","Medium",
 IF([@[RoastType]]="L","Light",
 IF([@[RoastType]]="D","Dark",""))) 
Why: Improves readability in the UI (legends, slicers) so stakeholders don’t need a codebook.

### 4.3 Loyalty Card (Late-Added Field)
I added a Loyalty Card column to Orders and filled it via XLOOKUP from the Customers table.
Why (and why a Table matters): Since Orders was converted to an Excel Table, any new column automatically became part of the Pivot data source after a simple Refresh—no manual range resizing.

### 5) Data Formatting 
Dates: Custom format dd-mmm-yyyy (e.g., 05-Sep-2024)
Why: Month in text avoids misinterpretation across locales (Haivng different data formats in USA ans Europe).
Sizes: Custom numeric format with unit 0.0" kg" so it displays 1.0 kg, 0.5 kg, etc.
Notes : Adds unit context directly in the cell; improves comprehension.

Currency: Formatted Unit Price and Sales as USD (no decimals in Pivots)inorder to have consistent financial presentation in charts and tables.

Duplicates: Used Data → Remove Duplicates on the Orders table to ensures accuracy before analysis and avoids double-counted sales.

Convert range into Table and named it Orders.
Why: Tables auto-expand, simplify formulas ([@Col]), and make Pivots refresh reliably.

## 6) Analysis Model & Visuals 
### 6.1 Total Sales Over Time by Coffee Type (Line Chart)
Pivot setup:
Rows: Order Date → Group by Years and Months
Columns: Coffee Type Name
Values: Sum of Sales (formatted with thousands separator, 0 decimals)

Chart: Line chart → removed field buttons, added USD Y-axis title, styled series by coffee type.
Why: Shows time-series trend as a high-value KPI; splitting by coffee type highlights product mix shifts.

### 6.2 Sales by Country (Bar Chart)
Pivot setup:
Rows: Country
Values: Sum of Sales (USD)
Sort: By value to show top country first (adjusted for bar orientation)
Chart: Bar with data labels outside, series borders for emphasis, consistent color theme.
Why: Quick geographic performance comparison; highlights “where to focus.”

### 6.3 Top 5 Customers (Bar Chart)
Pivot setup:
Rows: Customer Name
Values: Sum of Sales
Value Filters: Top 5 by Sum of Sales
Sort: By value
Chart: Bar chart, formatted consistently with the Country chart.

Why: Classic Pareto view—surfaces key accounts and retention/expansion targets.

### 7) Interactivity — Timeline & Slicers
## 7.1 Timeline (Order Date)
Inserted Timeline tied to Order Date.
Applied custom style (purple theme, white text).
Why: Makes date selection faster and more intuitive than filter dropdowns.

### 7.2 Slicers
Added slicers for: Roast Type Name, Size, Loyalty Card.
Applied custom styles and adjusted layouts (e.g., Size = 2×2, Roast = 1×3).
Why: Button-based filters are dashboard-friendly—easy for non-technical users.

### 7.3 Report Connections
Connected the Timeline and all Slicers to all three PivotTables via Report Connections.
Why: Keeps all visuals in sync—true dashboard behavior.

## Conclusion 
Working on this Coffee Sales Analysis gave a lot of hands-on learning. I started with raw data that looked scattered across different sheets, and step by step I shaped it into a clean, connected dataset. Using lookups (XLOOKUP and INDEX/MATCH) gave me the confidence to handle relational data inside Excel, just like how analysts work with joins in SQL.  

Building calculated columns like Sales and human-readable names taught me the value of keeping business logic clear and visible instead of hiding it in charts. Formatting choices—like fixing date styles, adding units to sizes, and cleaning duplicates—reminded me that small details go a long way in making reports reliable and easy to read.  

The biggest takeaway came when I built PivotTables and connected them with slicers and a timeline. Suddenly the dataset turned into an interactive dashboard where stakeholders could slice results by roast, size, or loyalty card in seconds. That’s when I realized how powerful Excel can be as a self-service analytics tool, not just for number crunching.  

By the end, I not only had a polished dashboard but also a story: how raw orders can become insights about seasonality, customer concentration, or product mix. This mirrors the real-world workflow of a data analyst—cleaning, modeling, analyzing, and finally communicating results in a way others can act on.

 










