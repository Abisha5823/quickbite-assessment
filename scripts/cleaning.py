from pathlib import Path
import pandas as pd
# Project folders and files
BASE_DIR = Path(__file__).resolve().parents[1]
INPUT_FILE = BASE_DIR / "data" / "QuickBite_Assessment.xlsx"
OUTPUT_DIR = BASE_DIR / "cleaned_data"
CLEANED_FILE = OUTPUT_DIR / "Quick_Bite_Data_Cleaning.xlsx"
COMBINED_FILE = OUTPUT_DIR / "Quick_Bite_Combined_Data.xlsx"
DUPLICATE_REVIEW_FILE = OUTPUT_DIR / "Duplicate_ID_Review.xlsx"
DUPLICATE_REVIEW = []

# 1. Load the six sheets from Excel
def load_data():
    if not INPUT_FILE.exists():
        raise FileNotFoundError("Input file not found: " + str(INPUT_FILE))
    print("Loading QuickBite workbook...")
    sheets = pd.read_excel(INPUT_FILE, sheet_name=None)
    required_sheets = [
        "orders",
        "customers",
        "restaurants",
        "areas",
        "delivery_partners",
        "feedback"
    ]
    missing_sheets = set(required_sheets) - set(sheets.keys())
    if missing_sheets:
        raise ValueError("Missing sheets: " + str(missing_sheets))
    data = {}
    for name in required_sheets:
        data[name] = sheets[name].copy()
    return data
    
# 2. Make column names consistent
def standardize_columns(df):
    df = df.copy()
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(r"[^a-z0-9]+", "_", regex=True)
        .str.strip("_")
    )
    return df

# 3. Remove extra spaces from text columns
def clean_text_values(df):
    df = df.copy()
    text_columns = df.select_dtypes(
        include=["object", "string"]
    ).columns
    for column in text_columns:
        df[column] = (
            df[column]
            .astype("string")
            .str.strip()
            .replace("", pd.NA)
        )
    return df
# 4. Remove exact duplicate rows
def remove_exact_duplicates(df, table_name):
    before = len(df)
    df = df.drop_duplicates().copy()
    removed = before - len(df)
    print(table_name, "- Removed exact duplicate rows:", removed)
    return df

# 5. Handle duplicate primary keys
def clean_primary_key(df, key_column, table_name):
    global DUPLICATE_REVIEW
    df = df.copy()
    # A primary key should not be empty.
    df = df.dropna(subset=[key_column])
    duplicate_rows = df[
        df.duplicated(subset=[key_column], keep=False)
    ].copy()
    if duplicate_rows.empty:
        return df
    # Save conflicting records 
    for key, group in duplicate_rows.groupby(key_column, sort=False):
        if len(group.drop_duplicates()) > 1:
            review = group.copy()
            review.insert(0, "table_name", table_name)
            DUPLICATE_REVIEW.append(review)
    # Keep the record that has more filled values.
    # If two records have the same completeness, keep the first one.
    df["_completeness"] = df.notna().sum(axis=1)
    df["_source_order"] = range(len(df))
    df = df.sort_values(
        by=["_completeness", "_source_order"],
        ascending=[False, True]
    )
    df = df.drop_duplicates(
        subset=[key_column],
        keep="first"
    )
    df = df.drop(
        columns=["_completeness", "_source_order"]
    )
    print(table_name, "- Primary key duplicates checked.")
    return df
# 6. Standardize category names
def standardize_categories(data):
    category_mappings = {
        "restaurants": {
            "cuisine": {
                "biriyani": "Biryani",
                "biryani": "Biryani",
                "n. indian": "North Indian",
                "north-indian": "North Indian",
                "north indian": "North Indian",
                "south-indian": "South Indian",
                "south indian": "South Indian",
                "dessert": "Desserts",
                "desserts": "Desserts"
            }
        },
        "delivery_partners": {
            "vehicle_type": {
                "bike": "Bike",
                "motor bike": "Bike",
                "motorbike": "Bike",
                "activa": "Scooter",
                "scooty": "Scooter",
                "scooter": "Scooter",
                "ev": "Electric Scooter",
                "ev scooter": "Electric Scooter",
                "e-scooter": "Electric Scooter",
                "bicycle": "Bicycle",
                "cycle": "Bicycle"
            }
        },
        "orders": {
            "payment_mode": {
                "card": "Card",
                "wallet": "Wallet",
                "upi": "UPI",
                "gpay": "Google Pay",
                "google pay": "Google Pay",
                "phonepe": "PhonePe",
                "cash": "Cash",
                "cod": "Cash on Delivery",
                "cash on delivery": "Cash on Delivery",
                "credit card": "Credit Card",
                "debit card": "Debit Card",
                "amazon pay": "Amazon Pay",
                "paytm wallet": "Paytm Wallet"
            },
            "order_status": {
                "delivered": "Delivered",
                "cancelled": "Cancelled",
                "canceled": "Cancelled",
                "pending": "Pending",
                "processing": "Processing",
                "completed": "Completed"
            }
        }
    }
    for table_name, columns in category_mappings.items():

        for column_name, mapping in columns.items():

            if column_name not in data[table_name].columns:
                continue

            data[table_name][column_name] = (
                data[table_name][column_name].apply(
                    lambda value: mapping.get(
                        str(value).strip().lower(),
                        str(value).strip()
                    ) if pd.notna(value) else value
                )
            )
            print("Standardized", table_name + "." + column_name)

    return data
# 7. Check the category values after standardization
def verify_categories(data):

    columns_to_check = {
        "restaurants": ["cuisine"],
        "delivery_partners": ["vehicle_type"],
        "orders": ["payment_mode", "order_status"]
    }
    print("\n CATEGORY VERIFICATION")
    for table_name, columns in columns_to_check.items():
        for column in columns:
            print("\n" + table_name + "." + column)
            print(
                data[table_name][column]
                .value_counts(dropna=False)
                .to_string()
            )
    print("\nCategory checking completed.")
# 8. Convert date and numeric columns
def convert_data_types(data):
    date_columns = {
        "areas": [],
        "restaurants": ["onboard_date"],
        "delivery_partners": ["join_date"],
        "customers": ["signup_date"],
        "orders": ["order_datetime"],
        "feedback": ["feedback_date"]
    }
    numeric_columns = {
        "restaurants": ["rating"],
        "orders": ["order_amount", "delivery_minutes"],
        "feedback": ["rating"]
    }
    # Convert date columns.
    for table_name, columns in date_columns.items():
        for column in columns:
            original = data[table_name][column].copy()
            converted = pd.to_datetime(
                original,
                format="mixed",
                dayfirst=True,
                errors="coerce"
            )
            invalid_dates = original[
                original.notna() & converted.isna()
            ]
            if not invalid_dates.empty:
                print("\nInvalid dates in", table_name + "." + column)
                print(invalid_dates.unique()[:20])
            data[table_name][column] = converted
            print(
                table_name + "." + column,
                "- Missing dates:",
                converted.isna().sum()
            )
    # Convert numeric columns.
    for table_name, columns in numeric_columns.items():
        for column in columns:
            original = data[table_name][column].copy()
            if column == "order_amount":
                # Remove currency names and commas.
                cleaned = (
                    original.astype("string")
                    .str.replace(
                        r"(?i)\b(?:rs\.?|inr)\s*",
                        "",
                        regex=True
                    )
                    .str.replace(",", "", regex=False)
                    .str.strip()
                )
                converted = pd.to_numeric(
                    cleaned,
                    errors="coerce"
                )
            elif column == "delivery_minutes":
                # Get the number from values like "18 min" or "26m".
                cleaned = original.astype("string").str.extract(
                    r"(-?\d+(?:\.\d+)?)",
                    expand=False
                )
                converted = pd.to_numeric(
                    cleaned,
                    errors="coerce"
                )
            else:
                converted = pd.to_numeric(
                    original,
                    errors="coerce"
                )
            invalid_values = original[
                original.notna() & converted.isna()
            ]
            if not invalid_values.empty:
                print("\nInvalid values in", table_name + "." + column)
                print(invalid_values.unique()[:20])
            data[table_name][column] = converted
            print(
                table_name + "." + column,
                "- Missing values:",
                converted.isna().sum()
            )
    return data

# 9. Remove rows with missing mandatory values
def remove_missing_mandatory_values(data):

    mandatory_columns = {
        "areas": ["area_id", "area_name", "city"],
        "restaurants": [
            "restaurant_id",
            "restaurant_name",
            "area_id",
            "cuisine"
        ],
        "delivery_partners": [
            "partner_id",
            "partner_name"
        ],
        "customers": [
            "customer_id",
            "customer_name"
        ],
        "orders": [
            "order_id",
            "customer_id",
            "restaurant_id",
            "order_datetime",
            "order_amount",
            "order_status"
        ],
        "feedback": ["feedback_id", "order_id"]
    }
    for table_name, columns in mandatory_columns.items():
        print("\nChecking", table_name)
        print("Missing values:")
        print(data[table_name][columns].isna().sum())
        missing_rows = data[table_name][columns].isna().any(axis=1).sum()
        print("Rows with missing mandatory values:", missing_rows)
        before = len(data[table_name])
        data[table_name] = data[table_name].dropna(
            subset=columns
        ).copy()
        removed = before - len(data[table_name])
        print("Rows removed:", removed)
    return data

# 10. Check the relationships between tables
def validate_relationships(data):
    valid_area_ids = set(data["areas"]["area_id"])
    # Remove restaurants linked to areas that do not exist.
    data["restaurants"] = data["restaurants"][
        data["restaurants"]["area_id"].isin(valid_area_ids)
    ].copy()
    # Customer area is optional, so invalid values are made missing.
    data["customers"].loc[
        ~data["customers"]["area_id"].isin(valid_area_ids),
        "area_id"
    ] = pd.NA
    valid_customer_ids = set(data["customers"]["customer_id"])
    valid_restaurant_ids = set(data["restaurants"]["restaurant_id"])
    valid_partner_ids = set(data["delivery_partners"]["partner_id"])
    # Orders must have a valid customer and restaurant.
    data["orders"] = data["orders"][
        data["orders"]["customer_id"].isin(valid_customer_ids)
        &
        data["orders"]["restaurant_id"].isin(valid_restaurant_ids)
    ].copy()
    # Partner details can be missing for some orders.
    data["orders"].loc[
        ~data["orders"]["partner_id"].isin(valid_partner_ids),
        "partner_id"
    ] = pd.NA
    # Feedback must belong to an existing order.
    valid_order_ids = set(data["orders"]["order_id"])
    data["feedback"] = data["feedback"][
        data["feedback"]["order_id"].isin(valid_order_ids)
    ].copy()
    print("\nForeign key checking completed.")
    return data
# 11. Check primary keys after cleaning
def validate_primary_keys(data):
    primary_keys = {
        "areas": "area_id",
        "restaurants": "restaurant_id",
        "delivery_partners": "partner_id",
        "customers": "customer_id",
        "orders": "order_id",
        "feedback": "feedback_id"
    }
    for table_name, key in primary_keys.items():
        if data[table_name][key].isna().any():
            raise ValueError(table_name + " has missing primary keys.")
        if data[table_name][key].duplicated().any():
            raise ValueError(
                table_name + " has duplicate values in " + key
            )
    print("All primary keys are valid.")
# 12. Save the six cleaned tables into one Excel workbook
def export_cleaned_data(data):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(
        CLEANED_FILE,
        engine="openpyxl"
    ) as writer:
        for table_name, df in data.items():

            df.to_excel(
                writer,
                sheet_name=table_name,
                index=False
            )
    print("\nCleaned workbook created:", CLEANED_FILE)
# 13. Combine feedback for each order
def aggregate_feedback(feedback):
    def combine_comments(values):
        comments = (
            values.dropna()
            .astype(str)
            .str.strip()
        )
        comments = comments[comments != ""]
        unique_comments = pd.unique(comments)
        if len(unique_comments) == 0:
            return pd.NA
        return " | ".join(unique_comments)
    feedback_summary = (
        feedback
        .groupby("order_id", as_index=False)
        .agg(
            rating=("rating", "mean"),
            comment=("comment", combine_comments)
        )
    )
    return feedback_summary
# 14. Join the six tables into one dataset
def create_combined_data(data):
    orders = data["orders"].copy()
    customers = data["customers"].copy()
    restaurants = data["restaurants"].copy()
    areas = data["areas"].copy()
    partners = data["delivery_partners"].copy()
    feedback = data["feedback"].copy()
    # First combine feedback so each order has only one feedback row.
    feedback_summary = aggregate_feedback(feedback)
    # Add customer names.
    combined = orders.merge(
        customers[["customer_id", "customer_name"]],
        on="customer_id",
        how="left",
        validate="many_to_one"
    )
    # Add restaurant details.
    combined = combined.merge(
        restaurants[
            ["restaurant_id", "restaurant_name", "area_id", "cuisine"]
        ],
        on="restaurant_id",
        how="left",
        validate="many_to_one"
    )
    # Add area and city details.
    combined = combined.merge(
        areas[["area_id", "area_name", "city"]],
        on="area_id",
        how="left",
        validate="many_to_one"
    )
    # Add delivery partner details.
    combined = combined.merge(
        partners[["partner_id", "partner_name", "vehicle_type"]],
        on="partner_id",
        how="left",
        validate="many_to_one"
    )
    # Add feedback details.
    combined = combined.merge(
        feedback_summary,
        on="order_id",
        how="left",
        validate="one_to_one"
    )
    # Split order datetime into date, time and weekday.
    combined["order_date"] = combined["order_datetime"].dt.date
    combined["order_time"] = (
        combined["order_datetime"].dt.strftime("%H:%M:%S")
    )
    combined["order_day"] = (
        combined["order_datetime"].dt.day_name()
    )
    # More than 45 minutes means the order was late.
    # Missing delivery time stays unknown.
    combined["is_late"] = pd.Series(
        pd.NA,
        index=combined.index,
        dtype="boolean"
    )
    known_duration = combined["delivery_minutes"].notna()
    combined.loc[known_duration, "is_late"] = (
        combined.loc[known_duration, "delivery_minutes"] > 45
    )
    # only the columns required for the assessment.
    required_columns = [
        "order_id",
        "order_date",
        "order_time",
        "order_day",
        "customer_name",
        "restaurant_name",
        "cuisine",
        "area_name",
        "city",
        "partner_name",
        "vehicle_type",
        "order_amount",
        "payment_mode",
        "coupon_code",
        "order_status",
        "delivery_minutes",
        "is_late",
        "rating",
        "comment"
    ]
    combined = combined[required_columns].copy()
    # Check that every order appears only once.
    if combined["order_id"].duplicated().any():
        raise ValueError("Duplicate order IDs found in combined data.")
    if len(combined) != orders["order_id"].nunique():
        raise ValueError("Combined row count does not match unique orders.")
    return combined
# 15. Save the combined dataset
def export_combined_data(combined):
    with pd.ExcelWriter(
        COMBINED_FILE,
        engine="openpyxl"
    ) as writer:
        combined.to_excel(
            writer,
            sheet_name="Combined Data",
            index=False
        )
    print("\nCombined workbook created:", COMBINED_FILE)
    print("Total combined rows:", len(combined))
    print("Unique orders:", combined["order_id"].nunique())
# 16. Save conflicting duplicate records for checking
def export_duplicate_review():
    if DUPLICATE_REVIEW:
        review_df = pd.concat(
            DUPLICATE_REVIEW,
            ignore_index=True
        )
        review_df.to_excel(
            DUPLICATE_REVIEW_FILE,
            index=False
        )
        print("Duplicate records saved for review:", DUPLICATE_REVIEW_FILE)
# 17. Run all cleaning steps
def main():
    data = load_data()
    # Standardize columns and text in every table.
    for table_name in data:
        data[table_name] = standardize_columns(data[table_name])
        data[table_name] = clean_text_values(data[table_name])
        # Standardize coupon codes before removing duplicates.
        if table_name == "orders" and "coupon_code" in data[table_name].columns:

            data[table_name]["coupon_code"] = (
                data[table_name]["coupon_code"]
                .astype("string")
                .str.strip()
                .str.upper()
            )
            print("Standardized coupon codes.")
        data[table_name] = remove_exact_duplicates(
            data[table_name],
            table_name
        )
    # Standardize categories and check the results.
    data = standardize_categories(data)
    verify_categories(data)
    # Convert dates and numbers.
    data = convert_data_types(data)
    # Remove rows with missing important values.
    data = remove_missing_mandatory_values(data)
    # Check duplicate primary keys in all tables.
    primary_keys = {
        "areas": "area_id",
        "restaurants": "restaurant_id",
        "delivery_partners": "partner_id",
        "customers": "customer_id",
        "orders": "order_id",
        "feedback": "feedback_id"
    }
    for table_name, key in primary_keys.items():
        data[table_name] = clean_primary_key(
            data[table_name],
            key,
            table_name
        )
    # table relationships.
    data = validate_relationships(data)
    # primary keys one final time.
    validate_primary_keys(data)
    # Save cleaned tables.
    export_cleaned_data(data)
    # Combine all tables.
    combined = create_combined_data(data)
    # Save combined data.
    export_combined_data(combined)
    # Save duplicate records if any were found.
    export_duplicate_review()
    print("\n========== CLEANING COMPLETED ==========\n")
if __name__ == "__main__":
    main()
