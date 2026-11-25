QuickBite – Data Cleaning and Business Analysis
1. Understanding the Data
I started by opening the given Excel file and going through all six tables: orders, customers, restaurants, areas, delivery partners and feedback.

I checked the columns, sample data, data types, and IDs to understand what each table contained and how they were connected. I also looked at how the IDs, such as customer_id, restaurant_id, area_id, partner_id, and order_id, were used to link the tables.

This helped me understand the data and decide how to combine the tables for analysis.
2. Data Issues I Found and Fixed
Before starting the analysis, I checked the data for errors and inconsistencies.
- Duplicate records: I found repeated rows and duplicate IDs. I removed exact duplicates and checked duplicate IDs to avoid incorrect records.
- Column names and text: Some column names and text values had inconsistent formatting and extra spaces. I cleaned them and made the column names consistent.
- Inconsistent categories: Values such as cuisine, vehicle type, payment mode and order status had different formats. I standardised them so that the same category would not be counted separately.
- Date format issues: Some dates were in different formats or were not recognised properly. I converted them into a consistent datetime format and handled invalid dates. I also separated date and time wherever required.
- Numeric values stored as text: Some values, like 18 min and order amounts with currency symbols or commas, were stored as text. I removed the unnecessary characters and converted them into numeric values for calculations.
- Missing values: I checked missing values in all tables and handled them based on the column and its importance. I tried to retain useful records instead of removing rows unnecessarily.
- Invalid IDs: I checked whether the IDs in each table matched the related tables. I removed records with invalid mandatory IDs and handled invalid optional IDs as missing values.
3. Combining the Data
After cleaning the tables, I used Pandas merge() to combine them using their common IDs.

One issue I faced was that an order could have multiple feedback records. If I directly merged them, the same order could appear multiple times. To avoid this, I grouped the feedback by order_id, calculated the average rating and combined the available comments before merging.

The final combined dataset contained 1,406 orders and 19 columns, with one row for each order.
4. Business Analysis
I used the combined dataset to answer the three business questions. For the analysis, I considered only delivered orders.
Area Performance Analysis
I grouped the orders by area and calculated the order count, revenue, average order value, delivery time, late delivery rate and customer rating. I then compared the areas based on revenue.

Result: Madhapur recorded the highest revenue.
Top Cuisine by Day
I grouped the delivered orders by weekday and cuisine, counted the orders, and identified the most ordered cuisine for each day. I also calculated its percentage share of orders.

Result: I identified the most ordered cuisine for each weekday. For example, South Indian was the top cuisine on Monday and Tuesday, Chinese on Wednesday and Biryani on Friday.
Delivery Vehicle Performance
I grouped the orders by vehicle type and compared their delivery time, late delivery rate, on-time delivery rate, order value and customer rating.

Result: Bike deliveries had a lower late delivery rate, while bicycle deliveries had the highest late delivery rate among the vehicle types.

For late deliveries, I followed the given rule of considering deliveries taking more than 45 minutes as late. Orders with missing delivery time were treated as unknown and excluded from the late delivery rate calculation.
5. How to Run the Project

Install the required libraries:
pip install -r requirements.txt

Run the data cleaning script:
python scripts/cleaning.py

This generates the cleaned Excel workbook and combined dataset.

Then run the analysis script:
python scripts/analysis.py

This generates the final Excel workbook containing the results of all three business analyses.
6. Technologies Used
- Python
- Pandas
- NumPy
- OpenPyXL
