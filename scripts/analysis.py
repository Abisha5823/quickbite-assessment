from pathlib import Path
from copy import copy
import pandas as pd

# Project folders and files
BASE_DIR = Path(__file__).resolve().parents[1]
INPUT_FILE = BASE_DIR / "cleaned_data" / "Quick_Bite_Combined_Data.xlsx"
OUTPUT_DIR = BASE_DIR / "analysis"
OUTPUT_FILE = OUTPUT_DIR / "Quick_Bite_Analysis.xlsx"

WEEKDAY_ORDER = [
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"
]
# 1. Load the combined dataset
def load_combined_data():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            "Combined workbook not found: " + str(INPUT_FILE)
            + "\nRun cleaning.py before analysis.py."
        )
    df = pd.read_excel(INPUT_FILE)
    print("Combined dataset loaded.")
    print("Total rows:", len(df))
    print("Total columns:", len(df.columns))
    return df

# 2. Check required columns and order IDs
def validate_data(df):
    required_columns = [
        "order_id", "order_date", "cuisine", "area_name", "vehicle_type",
        "order_amount", "order_status", "delivery_minutes", "rating"
    ]
    missing_columns = set(required_columns) - set(df.columns)

    if missing_columns:
        raise ValueError(
            "Missing required columns: " + str(sorted(missing_columns))
        )
    if df.empty:
        raise ValueError("The combined dataset is empty.")

    if df["order_id"].isna().any():
        raise ValueError("Missing order IDs found.")
    if df["order_id"].duplicated().any():
        raise ValueError(
            "Duplicate order IDs found. There should be one row per order."
        )
    print("Required columns are available.")
    print("Order IDs are unique.")

# 3. Prepare the data for analysis
def prepare_analysis_data(df):
    df = df.copy()
    # Convert amount, delivery time and rating into numbers.
    numeric_columns = ["order_amount", "delivery_minutes", "rating"]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    # Convert order date into date format.
    df["order_date"] = pd.to_datetime(df["order_date"], errors="coerce")

    if df["order_date"].isna().any():
        raise ValueError("Missing or invalid order dates found.")

    # Make order status lowercase to avoid spelling differences.
    df["order_status"] = (
        df["order_status"]
        .astype("string")
        .str.strip()
        .str.casefold()
    )

    print("\nOrder status distribution:")
    print(df["order_status"].value_counts(dropna=False))

    # Only delivered orders are used for this analysis.
    delivered = df[df["order_status"] == "delivered"].copy()

    if delivered.empty:
        raise ValueError("No delivered orders found.")

    # Get the weekday from the order date.
    delivered["order_day"] = delivered["order_date"].dt.day_name()
    # More than 45 minutes is considered late.
    # Missing delivery time is kept as unknown.
    delivered["is_late"] = pd.Series(
        pd.NA, index=delivered.index, dtype="boolean"
    )
    known_duration = delivered["delivery_minutes"].notna()
    delivered.loc[known_duration, "is_late"] = (
        delivered.loc[known_duration, "delivery_minutes"] > 45
    )

    print("\nPREPARATION SUMMARY")
    print("Total delivered orders:", len(delivered))
    print("Orders with known delivery duration:", known_duration.sum())
    print("Orders with missing delivery duration:", (~known_duration).sum())
    print("Late orders:", int(delivered["is_late"].sum()))
    print("Orders with missing ratings:", delivered["rating"].isna().sum())
    print("Orders with missing amounts:", delivered["order_amount"].isna().sum())
    return delivered

# 4. Question 1 - Area Performance
def analyze_area_performance(df):
    print("\nAnalyzing area performance...")
    grouped = df.groupby("area_name", dropna=False)

    # Calculate the main performance values for each area.
    area_summary = grouped.agg(
        delivered_orders=("order_id", "nunique"),
        total_revenue=("order_amount", lambda values: values.sum(min_count=1)),
        orders_with_amount=("order_amount", "count"),
        average_delivery_minutes=("delivery_minutes", "mean"),
        known_delivery_count=("delivery_minutes", "count"),
        late_orders=("is_late", "sum"),
        average_rating=("rating", "mean")
    ).reset_index()
    # Count orders where delivery time is not available.
    area_summary["unknown_delivery_orders"] = (
        area_summary["delivered_orders"] - area_summary["known_delivery_count"]
    )
    # Average order value.
    area_summary["average_order_value"] = (
        area_summary["total_revenue"] / area_summary["orders_with_amount"]
    )
    # Percentage of late deliveries.
    area_summary["late_delivery_rate_pct"] = (
        area_summary["late_orders"] / area_summary["known_delivery_count"] * 100
    )
    # Number of orders delivered on time.
    area_summary["on_time_orders"] = (
        area_summary["known_delivery_count"] - area_summary["late_orders"]
    )
    # Percentage of on-time deliveries.
    area_summary["on_time_delivery_rate_pct"] = (
        area_summary["on_time_orders"] / area_summary["known_delivery_count"] * 100
    )
    # Rank areas based on revenue.
    area_summary["revenue_rank"] = (
        area_summary["total_revenue"]
        .rank(method="min", ascending=False)
        .astype("Int64")
    )
    # Highest revenue areas appear first.
    area_summary = area_summary.sort_values(
        by=["total_revenue", "area_name"],
        ascending=[False, True],
        na_position="last"
    )
    # Arrange the output columns.
    area_summary = area_summary[
        [
            "revenue_rank", "area_name", "delivered_orders", "total_revenue",
            "average_order_value", "average_delivery_minutes", "known_delivery_count",
            "unknown_delivery_orders", "late_orders", "late_delivery_rate_pct",
            "on_time_orders", "on_time_delivery_rate_pct", "average_rating"
        ]
    ]
    # Round values to two decimal places.
    numerical_columns = [
        "total_revenue", "average_order_value", "average_delivery_minutes",
        "late_delivery_rate_pct", "on_time_delivery_rate_pct", "average_rating"
    ]
    area_summary[numerical_columns] = area_summary[numerical_columns].round(2)
    return area_summary

# 5. Question 2 - Top Cuisine by Day
def analyze_top_cuisine_by_day(df):
    print("\nAnalyzing top cuisine by weekday...")

    # Find the total delivered orders for each weekday.
    weekday_totals = (
        df.groupby("order_day")["order_id"]
        .nunique()
        .rename("total_day_orders")
        .reset_index()
    )

    # Count orders for each cuisine on each weekday.
    cuisine_counts = (
        df.dropna(subset=["cuisine"])
        .groupby(["order_day", "cuisine"])["order_id"]
        .nunique()
        .reset_index(name="order_count")
    )

    # Add total weekday orders to the cuisine table.
    cuisine_counts = cuisine_counts.merge(
        weekday_totals, on="order_day", how="left", validate="many_to_one"
    )
    # Calculate the percentage of orders for each cuisine.
    cuisine_counts["order_share_pct"] = (
        cuisine_counts["order_count"] / cuisine_counts["total_day_orders"] * 100
    )
    # Find the highest order count for each weekday.
    cuisine_counts["maximum_orders"] = (
        cuisine_counts.groupby("order_day")["order_count"].transform("max")
    )
    # Keep the cuisine with the highest count.
    # If two cuisines have the same count, both are kept.
    top_cuisine = cuisine_counts[
        cuisine_counts["order_count"] == cuisine_counts["maximum_orders"]
    ].copy()
    # Arrange weekdays from Monday to Sunday.
    day_numbers = {day: index for index, day in enumerate(WEEKDAY_ORDER)}
    top_cuisine["weekday_number"] = top_cuisine["order_day"].map(day_numbers)

    top_cuisine = top_cuisine.sort_values(
        by=["weekday_number", "cuisine"]
    )
    # Keep only the required output columns.
    top_cuisine = top_cuisine[
        ["order_day", "cuisine", "order_count", "total_day_orders", "order_share_pct"]
    ]
    top_cuisine["order_share_pct"] = top_cuisine["order_share_pct"].round(2)
    return top_cuisine

# 6. Question 3 - Vehicle Performance
def analyze_vehicle_performance(df):
    print("\nAnalyzing vehicle performance...")
    df = df.copy()

    # Replace missing vehicle names with Unknown.
    df["vehicle_type"] = (
        df["vehicle_type"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .replace("", "Unknown")
    )

    grouped = df.groupby("vehicle_type", dropna=False)
    # Calculate performance values for each vehicle.
    vehicle_summary = grouped.agg(
        delivered_orders=("order_id", "nunique"),
        total_revenue=("order_amount", lambda values: values.sum(min_count=1)),
        orders_with_amount=("order_amount", "count"),
        average_delivery_minutes=("delivery_minutes", "mean"),
        known_delivery_count=("delivery_minutes", "count"),
        late_orders=("is_late", "sum"),
        average_rating=("rating", "mean")
    ).reset_index()

    # Count orders with missing delivery time.
    vehicle_summary["unknown_delivery_orders"] = (
        vehicle_summary["delivered_orders"] - vehicle_summary["known_delivery_count"]
    )
    # Calculate late delivery percentage.
    vehicle_summary["late_delivery_rate_pct"] = (
        vehicle_summary["late_orders"] / vehicle_summary["known_delivery_count"] * 100
    )
    # Calculate on-time deliveries.
    vehicle_summary["on_time_orders"] = (
        vehicle_summary["known_delivery_count"] - vehicle_summary["late_orders"]
    )
    # Calculate on-time delivery percentage.
    vehicle_summary["on_time_delivery_rate_pct"] = (
        vehicle_summary["on_time_orders"] / vehicle_summary["known_delivery_count"] * 100
    )
    # Calculate average order value.
    vehicle_summary["average_order_value"] = (
        vehicle_summary["total_revenue"] / vehicle_summary["orders_with_amount"]
    )
    # Lower late delivery percentage means better performance.
    vehicle_summary = vehicle_summary.sort_values(
        by="late_delivery_rate_pct", ascending=True, na_position="last"
    )
    # Arrange the output columns.
    vehicle_summary = vehicle_summary[
        [
            "vehicle_type", "delivered_orders", "total_revenue", "average_order_value",
            "average_delivery_minutes", "known_delivery_count", "unknown_delivery_orders",
            "late_orders", "late_delivery_rate_pct", "on_time_orders",
            "on_time_delivery_rate_pct", "average_rating"
        ]
    ]
    numerical_columns = [
        "total_revenue", "average_order_value", "average_delivery_minutes",
        "late_delivery_rate_pct", "on_time_delivery_rate_pct", "average_rating"
    ]
    vehicle_summary[numerical_columns] = vehicle_summary[numerical_columns].round(2)
    return vehicle_summary
# 7. Export all three analysis results to Excel
def export_analysis(area_performance, top_cuisine, vehicle_performance):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        area_performance.to_excel(writer, sheet_name="Area Performance", index=False)
        top_cuisine.to_excel(writer, sheet_name="Top Cuisine by Day", index=False)
        vehicle_performance.to_excel(writer, sheet_name="Vehicle Performance", index=False)

        # Format the Excel sheets so they are easier to read.
        for worksheet in writer.sheets.values():
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            # Adjust column widths based on the cell contents.
            for column_cells in worksheet.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                column_letter = column_cells[0].column_letter
                worksheet.column_dimensions[column_letter].width = min(
                    max(max_length + 2, 12), 32
                )

            # Make the first row bold.
            for cell in worksheet[1]:
                new_font = copy(cell.font)
                new_font.bold = True
                cell.font = new_font

    print("\nAnalysis workbook created.")
    print("Output file:", OUTPUT_FILE)

# 8. Run the complete analysis
def main():
    print("\nQUICKBITE ANALYSIS\n")

    # Load the combined data.
    df = load_combined_data()

    # Check the data before analysis.
    validate_data(df)

    # Keep only delivered orders and prepare the columns.
    delivered = prepare_analysis_data(df)

    # Question 1: Area performance.
    area_performance = analyze_area_performance(delivered)

    # Question 2: Top cuisine for each weekday.
    top_cuisine = analyze_top_cuisine_by_day(delivered)

    # Question 3: Vehicle performance.
    vehicle_performance = analyze_vehicle_performance(delivered)

    # Export all results.
    export_analysis(area_performance, top_cuisine, vehicle_performance)
    print("\nANALYSIS COMPLETED\n")

if __name__ == "__main__":
    main()
