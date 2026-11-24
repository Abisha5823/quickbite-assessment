
from pathlib import Path
import pandas as pd
from copy import copy


# ============================================================
# 1. PROJECT PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parents[1]

INPUT_FILE = (
    BASE_DIR
    / "cleaned_data"
    / "Quick_Bite_Combined_Data.xlsx"
)

OUTPUT_DIR = BASE_DIR / "analysis"

OUTPUT_FILE = (
    OUTPUT_DIR
    / "Quick_Bite_Analysis.xlsx"
)

WEEKDAY_ORDER = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
]


# ============================================================
# 2. LOAD COMBINED DATA
# ============================================================

def load_combined_data():

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input workbook not found: {INPUT_FILE}\n"
            "Run cleaning.py before analysis.py."
        )

    df = pd.read_excel(INPUT_FILE)

    print("Combined dataset loaded successfully.")
    print("Total rows:", len(df))
    print("Total columns:", len(df.columns))

    return df


# ============================================================
# 3. VALIDATE REQUIRED COLUMNS AND DATA QUALITY
# ============================================================

def validate_data(df):

    required_columns = [
        "order_id",
        "order_date",
        "cuisine",
        "area_name",
        "vehicle_type",
        "order_amount",
        "order_status",
        "delivery_minutes",
        "rating"
    ]

    missing_columns = (
        set(required_columns) - set(df.columns)
    )

    if missing_columns:
        raise ValueError(
            f"Missing required columns: {sorted(missing_columns)}"
        )

    if df["order_id"].isna().any():
        raise ValueError("Missing order IDs found.")

    if df["order_id"].duplicated().any():
        raise ValueError(
            "Duplicate order IDs found. "
            "Combined data must contain one row per order."
        )

    if df.empty:
        raise ValueError("The combined dataset is empty.")

    print("Required-column validation passed.")
    print("Order ID uniqueness validation passed.")


# ============================================================
# 4. PREPARE DATA FOR ANALYSIS
# ============================================================

def prepare_analysis_data(df):

    df = df.copy()

    # Convert numerical columns.
    numeric_columns = [
        "order_amount",
        "delivery_minutes",
        "rating"
    ]

    for column in numeric_columns:

        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        )

    # Convert dates.
    df["order_date"] = pd.to_datetime(
        df["order_date"],
        errors="coerce"
    )

    if df["order_date"].isna().any():
        raise ValueError(
            "Invalid or missing order dates found. "
            "Check the cleaning output before analysis."
        )

    # Standardize order status for filtering.
    df["order_status"] = (
        df["order_status"]
        .astype("string")
        .str.strip()
        .str.casefold()
    )

    # Show the status distribution.
    print("\nOrder status distribution:")
    print(df["order_status"].value_counts(dropna=False))

    # Analyze only delivered orders.
    delivered = df[
        df["order_status"] == "delivered"
    ].copy()

    if delivered.empty:
        raise ValueError("No delivered orders found.")

    # Derive weekday from actual order date.
    delivered["order_day"] = (
        delivered["order_date"].dt.day_name()
    )

    # Calculate late status.
    # More than 45 minutes = Late.
    # Missing duration = Unknown.

    delivered["is_late"] = pd.Series(
        pd.NA,
        index=delivered.index,
        dtype="boolean"
    )

    known_duration = (
        delivered["delivery_minutes"].notna()
    )

    delivered.loc[
        known_duration,
        "is_late"
    ] = (
        delivered.loc[
            known_duration,
            "delivery_minutes"
        ] > 45
    )

    print("\n========== PREPARATION SUMMARY ==========")

    print("Total delivered orders:", len(delivered))

    print(
        "Orders with known delivery duration:",
        known_duration.sum()
    )

    print(
        "Orders with missing delivery duration:",
        (~known_duration).sum()
    )

    print(
        "Late orders:",
        int(delivered["is_late"].sum())
    )

    print(
        "Orders with missing ratings:",
        delivered["rating"].isna().sum()
    )

    print(
        "Orders with missing amounts:",
        delivered["order_amount"].isna().sum()
    )

    return delivered


# ============================================================
# 5. QUESTION 1: AREA PERFORMANCE
# ============================================================

def analyze_area_performance(df):

    print("\nAnalyzing area performance...")

    grouped = df.groupby(
        "area_name",
        dropna=False
    )

    area_summary = grouped.agg(

        delivered_orders=(
            "order_id",
            "nunique"
        ),

        total_revenue=(
            "order_amount",
            lambda values: values.sum(min_count=1)
        ),

        orders_with_amount=(
            "order_amount",
            "count"
        ),

        average_delivery_minutes=(
            "delivery_minutes",
            "mean"
        ),

        known_delivery_count=(
            "delivery_minutes",
            "count"
        ),

        late_orders=(
            "is_late",
            "sum"
        ),

        average_rating=(
            "rating",
            "mean"
        )

    ).reset_index()

    # Number of orders without delivery duration.
    area_summary["unknown_delivery_orders"] = (
        area_summary["delivered_orders"]
        - area_summary["known_delivery_count"]
    )

    # Average order value uses orders with known amounts.
    area_summary["average_order_value"] = (
        area_summary["total_revenue"]
        / area_summary["orders_with_amount"]
    )

    # Late delivery rate.
    area_summary["late_delivery_rate_pct"] = (
        area_summary["late_orders"]
        / area_summary["known_delivery_count"]
        * 100
    )

    # On-time delivery count.
    area_summary["on_time_orders"] = (
        area_summary["known_delivery_count"]
        - area_summary["late_orders"]
    )

    # On-time delivery percentage.
    area_summary["on_time_delivery_rate_pct"] = (
        area_summary["on_time_orders"]
        / area_summary["known_delivery_count"]
        * 100
    )

    # Rank areas by revenue.
    area_summary["revenue_rank"] = (
        area_summary["total_revenue"]
        .rank(method="min", ascending=False)
        .astype("Int64")
    )

    area_summary = area_summary.sort_values(
        by=["total_revenue", "area_name"],
        ascending=[False, True],
        na_position="last"
    )

    # Arrange columns.
    area_summary = area_summary[
        [
            "revenue_rank",
            "area_name",
            "delivered_orders",
            "total_revenue",
            "average_order_value",
            "average_delivery_minutes",
            "known_delivery_count",
            "unknown_delivery_orders",
            "late_orders",
            "late_delivery_rate_pct",
            "on_time_orders",
            "on_time_delivery_rate_pct",
            "average_rating"
        ]
    ]

    numerical_columns = [
        "total_revenue",
        "average_order_value",
        "average_delivery_minutes",
        "late_delivery_rate_pct",
        "on_time_delivery_rate_pct",
        "average_rating"
    ]

    area_summary[numerical_columns] = (
        area_summary[numerical_columns].round(2)
    )

    return area_summary


# ============================================================
# 6. QUESTION 2: TOP CUISINE BY DAY
# ============================================================

def analyze_top_cuisine_by_day(df):

    print("\nAnalyzing top cuisine by weekday...")

    # Total delivered orders for each weekday.
    weekday_totals = (
        df.groupby("order_day")["order_id"]
        .nunique()
        .rename("total_day_orders")
        .reset_index()
    )

    # Count orders for each cuisine on each weekday.
    cuisine_counts = (
        df.dropna(subset=["cuisine"])
        .groupby(
            ["order_day", "cuisine"]
        )["order_id"]
        .nunique()
        .reset_index(name="order_count")
    )

    # Add weekday totals.
    cuisine_counts = cuisine_counts.merge(
        weekday_totals,
        on="order_day",
        how="left",
        validate="many_to_one"
    )

    # Calculate cuisine share against all delivered
    # orders on that weekday.
    cuisine_counts["order_share_pct"] = (
        cuisine_counts["order_count"]
        / cuisine_counts["total_day_orders"]
        * 100
    )

    # Find the highest cuisine order count per weekday.
    cuisine_counts["maximum_orders"] = (
        cuisine_counts
        .groupby("order_day")["order_count"]
        .transform("max")
    )

    # Retain all top cuisines in case of a tie.
    top_cuisine = cuisine_counts[
        cuisine_counts["order_count"]
        == cuisine_counts["maximum_orders"]
    ].copy()

    # Sort weekdays in calendar order.
    top_cuisine["weekday_number"] = (
        top_cuisine["order_day"].map(
            {
                day: index
                for index, day in enumerate(WEEKDAY_ORDER)
            }
        )
    )

    top_cuisine = top_cuisine.sort_values(
        by=["weekday_number", "cuisine"]
    )

    top_cuisine = top_cuisine[
        [
            "order_day",
            "cuisine",
            "order_count",
            "total_day_orders",
            "order_share_pct"
        ]
    ]

    top_cuisine["order_share_pct"] = (
        top_cuisine["order_share_pct"].round(2)
    )

    return top_cuisine


# ============================================================
# 7. QUESTION 3: VEHICLE PERFORMANCE
# ============================================================

def analyze_vehicle_performance(df):

    print("\nAnalyzing vehicle performance...")

    df = df.copy()

    # Keep missing vehicle details visible.
    df["vehicle_type"] = (
        df["vehicle_type"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .replace("", "Unknown")
    )

    grouped = df.groupby(
        "vehicle_type",
        dropna=False
    )

    vehicle_summary = grouped.agg(

        delivered_orders=(
            "order_id",
            "nunique"
        ),

        total_revenue=(
            "order_amount",
            lambda values: values.sum(min_count=1)
        ),

        orders_with_amount=(
            "order_amount",
            "count"
        ),

        average_delivery_minutes=(
            "delivery_minutes",
            "mean"
        ),

        known_delivery_count=(
            "delivery_minutes",
            "count"
        ),

        late_orders=(
            "is_late",
            "sum"
        ),

        average_rating=(
            "rating",
            "mean"
        )

    ).reset_index()

    vehicle_summary["unknown_delivery_orders"] = (
        vehicle_summary["delivered_orders"]
        - vehicle_summary["known_delivery_count"]
    )

    # Calculate late delivery percentage.
    vehicle_summary["late_delivery_rate_pct"] = (
        vehicle_summary["late_orders"]
        / vehicle_summary["known_delivery_count"]
        * 100
    )

    # Calculate on-time deliveries.
    vehicle_summary["on_time_orders"] = (
        vehicle_summary["known_delivery_count"]
        - vehicle_summary["late_orders"]
    )

    # Calculate on-time delivery percentage.
    vehicle_summary["on_time_delivery_rate_pct"] = (
        vehicle_summary["on_time_orders"]
        / vehicle_summary["known_delivery_count"]
        * 100
    )

    # Calculate average order value.
    vehicle_summary["average_order_value"] = (
        vehicle_summary["total_revenue"]
        / vehicle_summary["orders_with_amount"]
    )

    # Sort by late delivery rate.
    # Lower rate means better delivery performance.
    vehicle_summary = vehicle_summary.sort_values(
        by="late_delivery_rate_pct",
        ascending=True,
        na_position="last"
    )

    vehicle_summary = vehicle_summary[
        [
            "vehicle_type",
            "delivered_orders",
            "total_revenue",
            "average_order_value",
            "average_delivery_minutes",
            "known_delivery_count",
            "unknown_delivery_orders",
            "late_orders",
            "late_delivery_rate_pct",
            "on_time_orders",
            "on_time_delivery_rate_pct",
            "average_rating"
        ]
    ]

    numerical_columns = [
        "total_revenue",
        "average_order_value",
        "average_delivery_minutes",
        "late_delivery_rate_pct",
        "on_time_delivery_rate_pct",
        "average_rating"
    ]

    vehicle_summary[numerical_columns] = (
        vehicle_summary[numerical_columns].round(2)
    )

    return vehicle_summary


# ============================================================
# 8. EXPORT ANALYSIS TO EXCEL
# ============================================================

def export_analysis(
    area_performance,
    top_cuisine,
    vehicle_performance
):

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        area_performance.to_excel(
            writer,
            sheet_name="Area Performance",
            index=False
        )

        top_cuisine.to_excel(
            writer,
            sheet_name="Top Cuisine by Day",
            index=False
        )

        vehicle_performance.to_excel(
            writer,
            sheet_name="Vehicle Performance",
            index=False
        )

        # Format all worksheets.
        for worksheet in writer.sheets.values():

            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

            # Make column widths readable.
            for column_cells in worksheet.columns:

                max_length = max(
                    len(str(cell.value))
                    if cell.value is not None else 0
                    for cell in column_cells
                )

                column_letter = (
                    column_cells[0].column_letter
                )

                worksheet.column_dimensions[
                    column_letter
                ].width = min(max(max_length + 2, 12), 32)

            # Format header row.
            for cell in worksheet[1]:
                new_font = copy(cell.font)
                new_font.bold = True
                cell.font = new_font

    print("\nAnalysis workbook exported successfully.")
    print("Output:", OUTPUT_FILE)


# ============================================================
# 9. MAIN EXECUTION
# ============================================================

def main():

    print("\n========== QUICKBITE ANALYSIS ==========\n")

    # Step 1: Load data.
    df = load_combined_data()

    # Step 2: Validate data.
    validate_data(df)

    # Step 3: Prepare delivered orders.
    delivered = prepare_analysis_data(df)

    # Step 4: Area performance.
    area_performance = analyze_area_performance(
        delivered
    )

    # Step 5: Top cuisine by weekday.
    top_cuisine = analyze_top_cuisine_by_day(
        delivered
    )

    # Step 6: Vehicle performance.
    vehicle_performance = analyze_vehicle_performance(
        delivered
    )

    # Step 7: Export results.
    export_analysis(
        area_performance,
        top_cuisine,
        vehicle_performance
    )

    print("\n========== ANALYSIS COMPLETED ==========\n")


if __name__ == "__main__":
    main()
