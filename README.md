# From-Raw-Data-to-Business-Decisions-A-Complete-Power-BI-Solution-for-JCars-Logistics
Data audit, currency standardisation, Power Query, star-schema modelling, DAX, dashboard design and evidence-based recommendations on a deliberately dirty Kenyan vehicle-sales dataset.
# Introduction
A vehicle sales report can look convincing and still mislead. A revenue total is only useful when its currency is clear, the underlying price fields agree, returns are handled consistently, and the report grain is understood. For JCars Logistics, those questions matter because a single flat file brings together vehicle sales, customers, branches, representatives, payments, delivery activity, costs, and customer feedback.
The goal of this article is to explain the analytical decisions behind that structure, on how to think about data quality, how to model the subject, how to define reusable calculations, and how to make findings responsibly. The dataset is small enough to inspect row by row, but its mixed formats and conflicting values make it a useful example of why preparation and validation are part of analysis, not housekeeping.
This project turns that operational dataset into a Power BI analysis intended to help management explore sales, profitability, operational performance, and areas needing investigation. The supplied report contains dedicated pages for calculations, the model, branches, sales representatives, payments, customers, and revenue. Its measure set includes total transactions, sales revenue, cost, gross profit, gross profit margin, revenue per car, units sold, discounted units, logistics cost ratio, logistics cost per car, and orders requiring investigation.

## Data Cleaning
I was working with a deliberately dirty flat file of 276 vehicle sales records from a Kenyan importer, JCars Logistics, and asked to turn it into a Power BI solution that helps management make decisions. This article walks through how I did it: understanding the grain, auditing data quality, cleaning with Power Query, modelling, writing DAX, building the report, and pulling out insights.

The most important thing I learned is that the headline numbers lied. As recorded, the business shows a 20.8% gross margin. Once I stopped trusting two suspicious transactions, the margin fell to 8.0%. I also found bugs in my own model while reconciling my numbers, and I document them here because they are the most useful part of the analysis.

### The Business Problem
JCars Logistics imports, sells, and delivers vehicles across Kenya. Management needs to understand not only how much business was recorded, but also where value is generated and what operational conditions may be associated with it. Questions include:

- How many transactions and vehicles are represented, and what revenue and gross profit do they produce?
- Which makes, models, branches, regions, and sales representatives contribute to performance?
- How do payment method and payment status relate to recorded sales?
- What does delivery status indicate about fulfilment, and where do logistics costs deserve attention?
- How do customer type, returns, and customer ratings add context to the sales picture?
- Which unusual records should be checked before a manager acts on a result?

### Understanding the Dataset
The data set had `276 rows` and `32 columns` in a single flat CSV.

**Grain**: One row is one sales order line, meaning one customer buying a specific vehicle (make, model, year, colour, transmission, fuel) in some quantity, at one branch, handled by one sales rep, with its own payment, delivery and rating information.

Column groups:

Group	Columns
**Transaction**	`Order ID, Order Date, Delivery Date`
**Customer**	`Customer Name, Customer Type, Customer Age`
**Geography**	`Region, County, City, Branch`
**People / channel**	`Sales Rep, Lead Source`
**Vehicle**	`Car Make, Car Model, Vehicle Type, Vehicle Year, Fuel Type, Transmission, Color`
**Measures**	`Units Sold, Unit Selling Price, Unit Cost, Discount, Delivery Fee, Logistics Cost, Revenue Recorded`
**Status**	`Payment Method, Payment Status, Delivery Status, Returned`
**Experience**	`Customer Rating, Review Count`


![Image description](https://dev-to-uploads.s3.us-east-2.amazonaws.com/uploads/articles/9e2wry2k0iq47wllfxw1.png)

### Data Quality Assessment
In the first few rows alone, dates use multiple representations, a delivery date appears as an Excel serial number; amounts include currency symbols, text suffixes, and plain numbers; categories vary in casing; and vehicle and employee names have spelling variants. The file also contains missing values in key fields, including Order ID, dates, customer details, unit cost, units sold, discount, logistics cost, customer rating, and return status.
These are the challenges that I faced while was working on the data:
**Missing transaction identifiers.** Missing or repeated Order IDs weaken distinct-order counts and make reconciliation difficult. Preserve the record for analysis where possible, but do not silently treat it as a confidently identified order.
**Mixed date formats.** Text dates in different patterns and spreadsheet serial values can parse differently depending on locale. Convert using explicit rules, then compare order and delivery chronology and flag unresolved values.
**Inconsistent monetary formats.** Values such as `KSh 8,136,000`, `KES 56,000`, plain numbers, and abbreviated text such as `9.14M` required parsing into numeric amounts and currency codes before calculations.
**Currency ambiguity.** Kenya Shillings are the required reporting currency. I treated the values without any currency as KES.
**Conflicting revenue fields.** Unit selling price, units sold, discount, and Revenue Recorded permit an independent revenue check. 
**Negative or implausible monetary values.** A negative recorded revenue may represent a return, a correction, or a mistake. 
**Category spelling and casing.** Examples in the source include `Toyta` versus `Toyota`, `petrol` versus `Petrol`, and multiple spellings/casings for branches and locations.
**Nonstandard categorical values.** Payment statuses include variants such as `Paid` and `complete`; delivery statuses include `AT YARD`, `held`, and `Delayed`.
**Inconsistent numeric types.** Unit costs, prices, discounts, and ratings appear as text in some rows. Parse percentages and numeric amounts separately, and capture parse failures instead of converting them to zero.
**Missing unit cost and units sold.** These fields underpin gross profit and unit metrics. A missing cost should not silently become zero, because that would overstate profit; a missing quantity should not be assumed to equal one without a documented rule.
**Unverified duplicate orders.** Exact duplicate rows and repeated identifiers can inflate revenue and volume. I investigated using transaction keys and row attributes.

### Business Definitoions Calculations
**Total Sales Revenue** = `SUMX('Jcars_Fact Table', 'Jcars_Fact Table'[Units Sold] * 'Jcars_Fact Table'[Unit Selling Price] * (1-COALESCE('Jcars_Fact Table'[Discount], 0)) + SUM('Jcars_Fact Table'[Delivery Fee]))`

![Image description](https://dev-to-uploads.s3.us-east-2.amazonaws.com/uploads/articles/lrltvvu3vve0r4mg4zne.png)

**Total Cost** = `SUMX('Jcars_Fact Table', 'Jcars_Fact Table'[Units Sold]* 'Jcars_Fact Table'[Unit Cost])`

**Total Logistics** = `SUM('Jcars_Fact Table'[Logistics Cost])`

**Total Transactions** = `DISTINCTCOUNT('Jcars_Fact Table'[Order ID])`

**Total Units Sold** = `SUM('Jcars_Fact Table'[Units Sold])`

**Rated Transactions** = 
`CALCULATE(
    DISTINCTCOUNT('Jcars_Fact Table'[Order ID]),
    NOT ISBLANK('Jcars_Fact Table'[Customer Rating]))`

**Discounted Revenue** = `CALCULATE([Total Sales Revenue], 'Jcars_Fact Table'[Discount]>0)`

## Data Modelling
The flat table repeats every descriptive attribute on every row, mixes measures with attributes, and cannot support reliable time intelligence. A star schema gives one narrow fact table and small, clean dimensions, so filters flow predictably.

![Image description](https://dev-to-uploads.s3.us-east-2.amazonaws.com/uploads/articles/tz0x2d85gji2m931qxgu.png)

I modelled my data into the following dimension tables:
`Dim_Customer`    One customer profile
`Dim_Location`    One branch (region, county, city, branch)
`Dim_SalesRep`    One salesrep
`Dim_Payment`     One method and status combination
`Dim_Vehicle`     One make, model, type, year, fuel, transmission, colour combination

`Fact_Table `holds the surrogate key, the dimension keys, the measures (units, KES price, KES cost, discount, KES delivery fee, KES logistics cost), delivery status, returned flag, rating, review count, order and delivery dates, and the data quality flag. It is deliberately free of descriptive text.


![Image description](https://dev-to-uploads.s3.us-east-2.amazonaws.com/uploads/articles/elptoj241vy3h8umx9pw.png)

## Analysis and Answers
**Which parts of the business have high logistics costs relative to value?** Overall logistics is 1.6% of revenue. The pressure lies in cheaper vehicles and specific places:

![Image description](https://dev-to-uploads.s3.us-east-2.amazonaws.com/uploads/articles/cjnk8ydirtattsfjd7s3.png)

Models with at least 8 orders: Corolla 4.5% of revenue (KES 83,600 per car), Demio 4.5%, Fielder 3.5%, Fit 3.0%, Outlander 3.0%, Canter 2.8%.
Vehicle types: hatchbacks and wagons 3.1%, crossovers 2.7%. Per car, trucks (KES 81,600), wagons (KES 73,600) and pickups (KES 70,400) cost the most to move.
Locations per car: Nairobi HQ KES 78,300, Mombasa KES 74,900, Athi River KES 72,300, against Nakuru KES 46,900 and Rift Valley KES 51,800.
36 individual orders spend more than 5% of their value on logistics. They are 4.6% of revenue but 21% of logistics cost. The worst are single low-value cars such as a Fit (14.6%), a Demio (12.3%) and an Axela (12.0%).

**Which customer types contribute most?**
Dealers are the most valuable segment: largest, second-best margin, cheapest to serve. NGOs generate the most orders (58) but earn the lowest margin. Individuals cost the most to deliver per car


![Image description](https://dev-to-uploads.s3.us-east-2.amazonaws.com/uploads/articles/rd7attttqlsxg0mwibz5.png)

**What is the relationship between discounts and performance?
**
The correlation between discount and margin is `-0.85.`
Discounts show almost no link to volume (-0.11 with cars per order, -0.12 with revenue).
Since the average price-to-cost ratio is 1.20, the break-even discount is about 16%. 21 orders sell below cost, together losing KES 8.9M, and all carry a discount of at least 5% (average 21%).
Orders discounted more than 10% are 23% of all orders and KES 297M of revenue, but earn only KES 2.0M of profit: a margin of 0.7%, against 14.0% for the rest.

**Which orders or customers should management investigate?**
`(blank ID)`	Harrier	Thika	(10.7M)	Delivery cancelled but partly paid, marked returned, 68-day gap, 15% discount
`LCL1095`	Impreza	Nairobi HQ	(1.6M)	Partly paid, delivery cancelled, marked returned, 15% discount
`LCL1195`	CX-5	Kakamega	(24.8M)	20% discount sold below cost, partly paid, delivery cancelled
`LCL1201`	Land Cruiser	Thika	(12.9M)	Paid, but delivery cancelled, date missing
`LCL1137`	Hilux	Nakuru	(2.5M)	50% discount, returned
`LCL1028`	Impreza	Kisumu	(1.0M)	50% discount, returned

## Findings and recommendations: what the data can support

The report design makes it possible to investigate where performance varies, but a defensible published article should only state results after checking the current model output under the intended filters. The files inspected for this write-up establish the source structure, report pages, measure names, and visible data-quality risks. They do not establish a validated headline revenue, profit, top branch, best-selling make, or delivery rate. Those figures should be read from the refreshed Power BI model after currency and revenue reconciliation, not inferred from a few sample records.
That distinction is important. A high-revenue branch may also have higher costs or a larger transaction base, a high logistics-cost ratio may be driven by vehicle mix or geography, and a low rating may have very few reviews. The report should show denominators and context, and the article should state when a pattern is descriptive rather than causal.

Based on the issues surfaced, management and the analyst should prioritize these evidence-led next steps:

**Reconcile recorded revenue before using it for targets or incentives.** Compare recorded revenue with quantity, unit selling price, discount, delivery fees, and returns. Investigate material mismatches and keep an exception list until the accounting rule is confirmed.
**Treat delivery and logistics differences as investigation signals.** Compare logistics cost per unit and delivery status by branch, region, and vehicle type. Review the contributing orders before changing routes, delivery commitments, or branch processes.
**Improve transaction capture for decision-critical fields.** Require valid order IDs, dates, quantities, currency, unit cost, and return/payment statuses at entry. These fields directly affect transaction counts, margins, time analysis, and exception monitoring.
**Interpret customer ratings with sample size.** Compare ratings alongside review counts and customer/vehicle context before proposing service interventions.
**Keep an auditable exception workflow.** Surface missing costs, negative revenue, unresolved currencies, date conflicts, and unusual amounts in a dedicated investigation view, then record their resolution rather than silently dropping them.

These are recommendations for analysis and control improvements, not causal conclusions about why a particular branch or employee performs as it does.

##Conclusion
The most valuable step in this project was to distrust the first number. The same 276 rows told three different profit stories depending on two decisions, whether foreign currencies were converted, and whether two impossible prices were believed. Only after resolving both, I could answer the real business questions, discounting drives profit, half of revenue is unpaid, and fulfilment problems sit at Nairobi HQ and Kisumu.


