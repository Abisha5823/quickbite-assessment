# QuickBite – Data Cleaning and Business Analysis

## Project Overview

QuickBite is a food delivery dataset analysis project. The dataset contains information about orders, customers, restaurants, delivery partners, areas, and customer feedback.

The main goal of this project is to clean the raw data, combine the related tables, and analyse the data to find useful business insights.

The project focuses on three main areas:
- Data cleaning
- Data combination
- Business analysis

## Dataset

The original Excel workbook contains six tables:

| Table | Description |
|---|---|
| orders | Order details, date, amount, status, and delivery duration |
| customers | Customer details and area |
| restaurants | Restaurant names, cuisine, and area |
| areas | Area and city information |
| delivery_partners | Delivery partner details and vehicle type |
| feedback | Customer ratings and comments |

### Table Relationships

The tables are connected using common ID columns:

- `customer_id` – connects orders with customers
- `restaurant_id` – connects orders with restaurants
- `area_id` – connects customers and restaurants with areas
- `partner_id` – connects orders with delivery partners
- `order_id` – connects orders with feedback

Understanding these relationships helped me combine the tables correctly.

## Project Workflow

The project was completed in the following steps:

1. Inspected all six tables.
2. Identified missing values, duplicates, inconsistent formats, and invalid records.
3. Cleaned and standardised the data.
4. Checked the relationships between the tables.
5. Combined the cleaned tables into one dataset.
6. Created additional columns required for analysis.
7. Analysed the data to answer the business questions.
8. Exported the cleaned data and analysis results into Excel files.

## Data Cleaning

I used Python and Pandas to clean all six tables.

### 1. Standardising Column Names

- Converted column names to lowercase.
- Removed extra spaces.
- Replaced spaces and special characters with underscores.

Example:

`Order DateTime` → `order_datetime`

### 2. Cleaning Text Values

- Removed unnecessary spaces.
- Converted empty strings into missing values.
- Standardised text formats wherever required.

### 3. Standardising Coupon Codes

Converted coupon codes to uppercase to maintain a consistent format.

Example:

`save50` → `SAVE50`

### 4. Removing Duplicate Records

Removed exact duplicate rows from all six tables.

| Table | Duplicate Rows Removed |
|---|---:|
| orders | 55 |
| customers | 14 |
| restaurants | 2 |
| areas | 1 |
| delivery_partners | 1 |
| feedback | 18 |

### 5. Handling Duplicate IDs

Checked for duplicate primary keys and compared the records.

- Retained the more complete record.
- If two records had the same level of completeness, retained the first record.
- Conflicting records were handled separately for review when required.

### 6. Standardising Categories

Standardised values in columns such as:

- Cuisine
- Vehicle type
- Payment mode
- Order status

This helped avoid treating the same category as different values during analysis.

### 7. Converting Data Types

Converted date and numeric columns into suitable formats.

- Date columns – order date, signup date, onboard date, partner join date, and feedback date.
- Numeric columns – order amount, delivery duration, restaurant rating, and customer feedback rating.

I also cleaned values such as `18 min` and converted them into numeric values. For order amounts, I removed currency text and unnecessary commas before converting them into numbers.

Invalid values were handled as missing values.

### 8. Handling Missing Values

Identified missing values in all tables and handled them based on the column's importance.

- Removed records missing mandatory information.
- Retained orders with missing delivery partner information.
- Retained orders with missing delivery duration.
- Retained feedback records with missing ratings.

This helped preserve useful data instead of removing complete records unnecessarily.

### 9. Validating Table Relationships

Checked whether the IDs in each table referred to valid records in the related tables.

For example:
- Orders should refer to existing customers.
- Orders should refer to existing restaurants.
- Restaurants should refer to valid areas.
- Feedback should refer to existing orders.

Invalid mandatory relationships were handled by removing the affected records. Optional invalid references, such as delivery partner IDs, were treated as missing.

## Combining the Tables

After cleaning the tables, I combined them using Pandas `merge()` operations and their common ID columns.

The feedback table required special handling because one order could have multiple feedback records.

Before merging feedback with orders:
- Grouped feedback by `order_id`.
- Calculated the average rating.
- Combined the available comments.

This helped maintain one row per order in the final dataset.

### Final Combined Dataset

The combined dataset contains **1,406 orders and 19 columns**.

The final dataset includes:

- Order ID
- Order date
- Order time
- Order day
- Customer name
- Restaurant name
- Cuisine
- Area name
- City
- Delivery partner name
- Vehicle type
- Order amount
- Payment mode
- Coupon code
- Order status
- Delivery duration
- Late delivery status
- Rating
- Comment

I also checked that each order appeared only once in the final dataset.

## Derived Columns

I created additional columns to make the analysis easier.

| Column | Description |
|---|---|
| `order_date` | Date extracted from the original order datetime |
| `order_time` | Time extracted from the original order datetime |
| `order_day` | Weekday extracted from the order date |
| `is_late` | Identifies whether an order was late |

### Late Delivery Rule

The assessment defined a late delivery as an order taking more than 45 minutes.

- More than 45 minutes – Late
- 45 minutes or less – On-time
- Missing delivery duration – Unknown (`NA`)

Orders with missing delivery duration were not treated as on-time deliveries.

## Business Analysis

The analysis was performed using delivered orders only. Cancelled and other non-delivered orders were excluded.

### 1. Area Performance Analysis

The objective was to understand how different areas performed in terms of orders, revenue, and delivery.

The following metrics were calculated:

- Total delivered orders
- Total revenue
- Average order value
- Average delivery time
- Number of late deliveries
- Late delivery rate
- Number of on-time deliveries
- On-time delivery rate
- Average customer rating

**Late Delivery Rate:**

`Late Delivery Rate = (Late Orders / Orders with Known Delivery Duration) × 100`

Orders with unknown delivery duration were excluded from this calculation.

Areas were also ranked based on total revenue to identify the highest-revenue areas.

### 2. Top Cuisine by Day

The objective was to find the most popular cuisine on each weekday.

The analysis:
- Grouped orders by weekday and cuisine.
- Counted unique delivered orders.
- Identified the cuisine with the highest order count for each day.
- Calculated its percentage share of delivered orders for that day.

If two cuisines had the same highest order count, both were included.

### 3. Delivery Vehicle Performance

The objective was to compare delivery performance across different vehicle types.

The following metrics were calculated:

- Total delivered orders
- Average delivery time
- Number of late deliveries
- Late delivery rate
- Number of on-time deliveries
- On-time delivery rate
- Average order value
- Average customer rating

The vehicle types were compared based on their late delivery rates.

Orders without a vehicle type were grouped under `Unknown`.

## Key Insights

Based on the analysis:

- Madhapur recorded the highest revenue among the areas.
- South Indian cuisine was the most ordered cuisine on Monday and Tuesday.
- Chinese cuisine was the top cuisine on Wednesday.
- Biryani was the top cuisine on Friday.
- Pizza was the top cuisine on Saturday.
- North Indian cuisine was the top cuisine on Sunday.
- Bike deliveries had a lower late delivery rate than the other identified vehicle types.
- Bicycle deliveries had the highest late delivery rate among the identified vehicle types.

These insights were based on the cleaned dataset and delivered orders.

## Output Files

The project generates three Excel output files:

| File | Description |
|---|---|
| `Quick_Bite_Data_Cleaning.xlsx` | Contains all six cleaned tables |
| `Quick_Bite_Combined_Data.xlsx` | Contains the combined dataset with one row per order |
| `Quick_Bite_Analysis.xlsx` | Contains the results of the three business analyses |

## Project Structure

```text
quickbite-assessment/
│
├── data/
│   └── QuickBite_Assessment.xlsx
│
├── scripts/
│   ├── cleaning.py
│   └── analysis.py
│
├── cleaned_data/
│   ├── Quick_Bite_Data_Cleaning.xlsx
│   └── Quick_Bite_Combined_Data.xlsx
│
├── analysis/
│   └── Quick_Bite_Analysis.xlsx
│
├── README.md
└── requirements.txt
```

## Technologies Used

- **Python** – Used for the overall data cleaning and analysis.
- **Pandas** – Used for data cleaning, merging tables, grouping, calculations, and analysis.
- **NumPy** – Used for numerical operations.
- **OpenPyXL** – Used for reading and writing Excel files.

## How to Run the Project

### 1. Clone the Repository

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
```

### 2. Open the Project Folder

```bash
cd quickbite-assessment
```

### 3. Install the Required Libraries

```bash
pip install -r requirements.txt
```

### 4. Run the Data Cleaning Script

```bash
python scripts/cleaning.py
```

This creates the cleaned Excel workbook and the combined dataset.

### 5. Run the Analysis Script

```bash
python scripts/analysis.py
```

This reads the combined dataset and generates the final analysis workbook.

## What I Learned

Through this project, I learned how to approach a data analysis task step by step.

I gained experience in:
- Inspecting an unfamiliar dataset.
- Identifying data quality issues.
- Cleaning text, dates, and numeric values.
- Handling missing values and duplicates.
- Checking relationships between tables.
- Combining multiple tables using Pandas.
- Creating new columns for analysis.
- Grouping and calculating business metrics.
- Understanding business performance through data.

This project helped me understand the complete process of working with raw data, cleaning it, combining it, analysing it, and using the results to find meaningful business insights.