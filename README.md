Project Overview

Coffee Orders Analysis (Excel Analytics Project) 

This project is an end-to-end Excel data analytics project focused on analyzing coffee order sales performance. Using raw transactional data, the project demonstrates how Excel can be used to clean, enrich, analyze, and visualize business data in a structured and professional way. 

The goal was to transform raw order data into clear sales insights, identify top-selling products and customers, and present findings through interactive dashboards suitable for business decision-making.

-- Dataset & Workbook Structure

The Excel workbook is structured into multiple worksheets, each serving a specific analytical purpose:

Orders – Raw transactional data

Customers – Customer information (name, email, country, loyalty card)

Products – Coffee product attributes

Pivot Tables – Aggregated analysis

Charts – Visual representations of insights

Dashboard – Consolidated view of all key insights

--Data Cleaning, Transformation & Enrichment

1. Customer Data Enrichment (XLOOKUP)

Customer details were dynamically pulled into the orders table using XLOOKUP.

Formulas used:

=XLOOKUP(Customer_ID, Customers!A:A, Customers!B:B)
=XLOOKUP(Customer_ID, Customers!A:A, Customers!C:C)
=XLOOKUP(Customer_ID, Customers!A:A, Customers!D:D)


Fields populated:

Customer Name

Email

Country

(NOTE) This replicates table joins commonly done in SQL, but implemented in Excel.

2. Product Data Enrichment (INDEX + MATCH)

Product attributes were retrieved from the products table using INDEX and MATCH.

Formula pattern used:

=INDEX(Products!A:G,
       MATCH(Product_ID, Products!A:A, 0),
       MATCH(Header, Products!1:1, 0))


Fields populated:

Coffee Type

Roast Type

Size

Unit Price

📌 This approach allows flexible and scalable lookups across large datasets.

3. Sales Calculation; 

A new Sales column was created to calculate total revenue per order.

Formula used:

=Unit_Price * Quantity


📌 This metric is the foundation for all revenue-based analysis in the project.

4. Data Standardization & Categorization;

Coffee Type Name Expansion

Abbreviated coffee types were converted into full names using logical formulas, and stored in a new column.

Formula pattern used:

=IF(Type="Rob","Robusta",
 IF(Type="Exc","Excelsa",
 IF(Type="Ara","Arabica",
 IF(Type="Lib","Liberica",""))))

Roast Type Name Expansion

Roast abbreviations were converted into readable labels and stored in a new column.

Formula used:

=IF(Roast="M","Medium",
 IF(Roast="L","Light",
 IF(Roast="D","Dark","")))


📌 This improves ease of understanding and ensures accurate grouping in Pivot Tables.

5. Loyalty Card Identification;

A new Loyalty Card column was created using XLOOKUP to identify customers enrolled in the loyalty program.

Formula used:

=XLOOKUP(Customer_ID, Customers!A:A, Customers!Loyalty_Card_Column)

6. Data Formatting;

Several formatting steps were applied to improve clarity and consistency like;

Order date reformatted from 01/10/2019 → 01-Oct-2019

Coffee size reformatted from numeric values to include units (e.g. 0.5 → 0.5 Kg)

Unit Price and Sales columns formatted as currency ($)

Duplicate values checked and reviewed

Data range converted into an Excel Table for easier analysis and scalability

-- Pivot Tables & Visual Analysis

7. Pivot Table Creation

Pivot Tables were created to summarize and analyze sales performance efficiently.

Key metrics analyzed:

Total Sales

Sales by Coffee Type

Sales by Customer

Time-based sales trends

8. Pivot Charts & Interactivity

Multiple Pivot Charts were created, including;

Sales Trend Over Time 

Fields: Order Date, Coffee Type Name, Sum of Sales

Total Sales by Coffee Type

Identified best-selling coffee types

Top 5 Customers by Sales

Customers ranked by total purchase value

9. Interactive Filters;

It was created to enhance usability and exploration

A Timeline was added for date-based filtering

Slicers were created for:

Roast Type Name

Coffee Size

Loyalty Card Status

📌 This allows stakeholders to explore the data dynamically without modifying formulas.

📈 Key Highlights & Insights;

Clear identification of best-selling coffee types

Visibility into top-spending customers

Ability to analyze sales performance across different roast types and sizes

Loyalty card customers can be isolated to assess repeat-purchase behavior

Time-based analysis reveals sales trends and seasonality

-- Tools & Skills Demonstrated

Microsoft Excel

XLOOKUP

INDEX & MATCH

Logical formulas (IF)

Data cleaning & formatting

Pivot Tables & Pivot Charts

Slicers & Timelines

Dashboard design

Business-focused analytical thinking

-- Business Value

This project demonstrates how Excel can be used to:

Turn raw sales data into decision-ready insights

Support product, sales, and customer analysis

Build interactive dashboards for non-technical stakeholders

Replicate SQL-style joins and aggregations in Excel


📬 I’m open to collaboration and discussions. Feel free to reach out to me via [LinkedIn](https://www.linkedin.com/in/adedeji-daramola-729250247/).

.
