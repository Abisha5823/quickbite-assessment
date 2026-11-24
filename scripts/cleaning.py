
from pathlib import Path

import numpy as np
import pandas as pd


# ============================================================
# 1. PROJECT PATHS AND CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parents[1]

INPUT_FILE = BASE_DIR / "data" / "QuickBite_Assessment.xlsx"

OUTPUT_DIR = BASE_DIR / "cleaned_data"

CLEANED_FILE = OUTPUT_DIR / "Quick_Bite_Data_Cleaning.xlsx"

COMBINED_FILE = OUTPUT_DIR / "Quick_Bite_Combined_Data.xlsx"

DUPLICATE_REVIEW_FILE = OUTPUT_DIR / "Duplicate_ID_Review.xlsx"


# ============================================================
# 2. LOAD ALL SIX TABLES
# ============================================================

def load_data():

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input file not found: {INPUT_FILE}"
        )

    print("Loading QuickBite workbook...")

    sheets = pd.read_excel(
        INPUT_FILE,
        sheet_name=None
    )

    required_sheets = [
        "orders",
        "customers",
        "restaurants",
        "areas",
        "delivery_partners",
        "feedback"
    ]

    missing_sheets = set(required_sheets) - set(sheets)

    if missing_sheets:
        raise ValueError(
            f"Missing sheets: {missing_sheets}"
        )

    return {
        name: sheets[name].copy()
        for name in required_sheets
    }


# ============================================================
# 3. STANDARDIZE COLUMN NAMES
# ============================================================

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


# ============================================================
# 4. STANDARDIZE TEXT VALUES
# ============================================================

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


# ============================================================
# 5. REMOVE EXACT DUPLICATES
# ============================================================

def remove_exact_duplicates(df, table_name):

    before = len(df)

    df = df.drop_duplicates().copy()

    removed = before - len(df)

    print(
        f"{table_name}: Removed {removed} exact duplicate rows."
    )

    return df


# ============================================================
# 6. HANDLE DUPLICATE PRIMARY KEYS
# ============================================================

DUPLICATE_REVIEW = []


def clean_primary_key(df, key_column, table_name):

    global DUPLICATE_REVIEW

    df = df.copy()

    # Remove records without a primary key.
    df = df.dropna(subset=[key_column])

    duplicated = df[
        df.duplicated(
            subset=[key_column],
            keep=False
        )
    ].copy()

    if duplicated.empty:
        return df

    # Identify groups containing conflicting records.
    for key, group in duplicated.groupby(
        key_column,
        sort=False
    ):

        if len(group.drop_duplicates()) > 1:

            review = group.copy()

            review.insert(
                0,
                "table_name",
                table_name
            )

            DUPLICATE_REVIEW.append(review)

    # Prefer the most complete record for each key.
    # If completeness is equal, retain the first source record.

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

    print(
        f"{table_name}: Primary key duplicates resolved."
    )

    return df


# ============================================================
# 7. STANDARDIZE CATEGORY VALUES
# ============================================================


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
                data[table_name][column_name]
                .apply(
                    lambda value: mapping.get(
                        str(value).strip().lower(),
                        str(value).strip()
                    ) if pd.notna(value) else value
                )
            )

            print(
                f"Standardized {table_name}.{column_name}"
            )

    return data

def verify_categories(data):

    columns_to_check = {
        "restaurants": ["cuisine"],
        "delivery_partners": ["vehicle_type"],
        "orders": ["payment_mode", "order_status"]
    }

    print("\n========== CATEGORY VERIFICATION ==========")

    for table_name, columns in columns_to_check.items():

        for column in columns:

            print(f"\n{table_name}.{column}")

            print(
                data[table_name][column]
                .value_counts(dropna=False)
                .to_string()
            )

    print("\n========== VERIFICATION COMPLETED ==========")




# ============================================================
# 8. CONVERT DATES AND NUMERIC VALUES
# ============================================================


def convert_data_types(data):

    # All date columns across the six tables.
    date_columns = {
        "areas": [],

        "restaurants": [
            "onboard_date"
        ],

        "delivery_partners": [
            "join_date"
        ],

        "customers": [
            "signup_date"
        ],

        "orders": [
            "order_datetime"
        ],

        "feedback": [
            "feedback_date"
        ]
    }

    # Numeric columns.
    numeric_columns = {
        "restaurants": [
            "rating"
        ],

        "orders": [
            "order_amount",
            "delivery_minutes"
        ],

        "feedback": [
            "rating"
        ]
    }

    # Convert mixed date formats.
    for table_name, columns in date_columns.items():

        for column in columns:

            original = data[table_name][column].copy()

            converted = pd.to_datetime(
                original,
                format="mixed",
                dayfirst=True,
                errors="coerce"
            )

            # Identify values that failed conversion.
            invalid_dates = original[
                original.notna() & converted.isna()
            ]

            if not invalid_dates.empty:

                print(
                    f"\nInvalid dates in {table_name}.{column}:"
                )

                print(invalid_dates.unique()[:20])

            data[table_name][column] = converted

            print(
                f"{table_name}.{column}: "
                f"{converted.isna().sum()} missing dates"
            )

    # Convert numeric columns.
    
    for table_name, columns in numeric_columns.items():
        for column in columns:
            original = data[table_name][column].copy()

            if column == "order_amount":
                # Remove currency labels and thousands separators.
                cleaned = (
                    original.astype("string")
                    .str.replace(r"(?i)\b(?:rs\.?|inr)\s*", "", regex=True)
                    .str.replace(",", "", regex=False)
                    .str.strip()
                )
                converted = pd.to_numeric(cleaned, errors="coerce")

            elif column == "delivery_minutes":
                # Extract the numeric duration from values like
                # "18 min", "26m", and "26 mins".
                cleaned = (
                    original.astype("string")
                    .str.extract(r"(-?\d+(?:\.\d+)?)", expand=False)
                )
                converted = pd.to_numeric(cleaned, errors="coerce")

            else:
                converted = pd.to_numeric(original, errors="coerce")

            invalid_values = original[
                original.notna() & converted.isna()
            ]

            if not invalid_values.empty:
                print(f"\nInvalid values in {table_name}.{column}:")
                print(invalid_values.unique()[:20])

            data[table_name][column] = converted

            print(
                f"{table_name}.{column}: "
                f"{converted.isna().sum()} missing values"
            )

    return data




# ============================================================
# 9. REMOVE ROWS WITH MISSING MANDATORY DATA
# ============================================================

def remove_missing_mandatory_values(data):

    mandatory_columns = {

        "areas": [
            "area_id",
            "area_name",
            "city"
        ],

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

        "feedback": [
            "feedback_id",
            "order_id"
        ]
    }

    for table_name, columns in mandatory_columns.items():

        print(f"\nChecking mandatory columns in: {table_name}")

        print("Missing values by column:")

        print(data[table_name][columns].isna().sum())

        print(
            "Rows having at least one missing mandatory value:",
            data[table_name][columns].isna().any(axis=1).sum()
        )

        before = len(data[table_name])

        data[table_name] = data[table_name].dropna(
            subset=columns
        ).copy()

        removed = before - len(data[table_name])

        print(
            f"{table_name}: Removed {removed} rows "
            "with missing mandatory values."
        )

    return data

# ============================================================
# 10. VALIDATE FOREIGN KEY RELATIONSHIPS
# ============================================================

def validate_relationships(data):

    # Areas referenced by restaurants.
    valid_area_ids = set(
        data["areas"]["area_id"]
    )

    # Remove restaurants whose area does not exist.
    data["restaurants"] = data["restaurants"][
        data["restaurants"]["area_id"].isin(valid_area_ids)
    ].copy()

    # Customer area is optional.
    # Invalid customer area references become missing.
    data["customers"].loc[
        ~data["customers"]["area_id"].isin(valid_area_ids),
        "area_id"
    ] = pd.NA

    valid_customer_ids = set(
        data["customers"]["customer_id"]
    )

    valid_restaurant_ids = set(
        data["restaurants"]["restaurant_id"]
    )

    valid_partner_ids = set(
        data["delivery_partners"]["partner_id"]
    )

    # Orders require valid customer and restaurant records.
    data["orders"] = data["orders"][
        data["orders"]["customer_id"].isin(valid_customer_ids)
        &
        data["orders"]["restaurant_id"].isin(valid_restaurant_ids)
    ].copy()

    # A delivery partner may be missing for cancelled
    # or otherwise undelivered orders.
    # Invalid partner references are set to missing.
    data["orders"].loc[
        ~data["orders"]["partner_id"].isin(valid_partner_ids),
        "partner_id"
    ] = pd.NA

    # Feedback must belong to an existing order.
    valid_order_ids = set(
        data["orders"]["order_id"]
    )

    data["feedback"] = data["feedback"][
        data["feedback"]["order_id"].isin(valid_order_ids)
    ].copy()

    print("Foreign key validation completed.")

    return data


# ============================================================
# 11. FINAL PRIMARY KEY VALIDATION
# ============================================================

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

            raise ValueError(
                f"{table_name} contains missing primary keys."
            )

        if data[table_name][key].duplicated().any():

            raise ValueError(
                f"{table_name} still contains duplicate {key} values."
            )

    print("All primary keys are valid.")


# ============================================================
# 12. EXPORT SIX CLEANED TABLES
# ============================================================

def export_cleaned_data(data):

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

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

    print("Cleaned workbook created:", CLEANED_FILE)


# ============================================================
# 13. AGGREGATE FEEDBACK
# ============================================================

def aggregate_feedback(feedback):

    def combine_comments(values):

        comments = (
            values
            .dropna()
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


# ============================================================
# 14. COMBINE ALL SIX TABLES
# ============================================================

def create_combined_data(data):

    orders = data["orders"].copy()

    customers = data["customers"].copy()

    restaurants = data["restaurants"].copy()

    areas = data["areas"].copy()

    partners = data["delivery_partners"].copy()

    feedback = data["feedback"].copy()

    # Aggregate feedback before joining.
    # This prevents multiple feedback records from
    # multiplying the number of order rows.

    feedback_summary = aggregate_feedback(feedback)

    # Join orders with customers.
    combined = orders.merge(
        customers[
            ["customer_id", "customer_name"]
        ],
        on="customer_id",
        how="left",
        validate="many_to_one"
    )

    # Join restaurants.
    combined = combined.merge(
        restaurants[
            [
                "restaurant_id",
                "restaurant_name",
                "area_id",
                "cuisine"
            ]
        ],
        on="restaurant_id",
        how="left",
        validate="many_to_one"
    )

    # Join area information.
    combined = combined.merge(
        areas[
            ["area_id", "area_name", "city"]
        ],
        on="area_id",
        how="left",
        validate="many_to_one"
    )

    # Join delivery partner information.
    combined = combined.merge(
        partners[
            ["partner_id", "partner_name", "vehicle_type"]
        ],
        on="partner_id",
        how="left",
        validate="many_to_one"
    )

    # Join aggregated feedback.
    combined = combined.merge(
        feedback_summary,
        on="order_id",
        how="left",
        validate="one_to_one"
    )

    # Derive date and time fields.
    combined["order_date"] = (
        combined["order_datetime"].dt.date
    )

    combined["order_time"] = (
        combined["order_datetime"].dt.strftime("%H:%M:%S")
    )

    combined["order_day"] = (
        combined["order_datetime"].dt.day_name()
    )

    # Late delivery rule:
    # More than 45 minutes = Late.
    # Missing delivery duration = Unknown.

    combined["is_late"] = pd.Series(
        pd.NA,
        index=combined.index,
        dtype="boolean"
    )

    known_duration = combined["delivery_minutes"].notna()

    combined.loc[
        known_duration,
        "is_late"
    ] = (
        combined.loc[
            known_duration,
            "delivery_minutes"
        ] > 45
    )

    # Keep the exact required output columns.
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

    # Ensure one row per order.
    if combined["order_id"].duplicated().any():

        raise ValueError(
            "Combined data contains duplicate order IDs."
        )

    if len(combined) != orders["order_id"].nunique():

        raise ValueError(
            "Combined data row count does not match unique orders."
        )

    return combined


# ============================================================
# 15. EXPORT COMBINED DATA
# ============================================================

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

    print("Combined workbook created:", COMBINED_FILE)

    print("Combined rows:", len(combined))

    print(
        "Unique orders:",
        combined["order_id"].nunique()
    )


# ============================================================
# 16. EXPORT DUPLICATE ID REVIEW
# ============================================================

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

        print(
            "Conflicting duplicate records saved to:",
            DUPLICATE_REVIEW_FILE
        )


# ============================================================
# 17. MAIN EXECUTION
# ============================================================

def main():

    print("\n========== QUICKBITE DATA CLEANING ==========\n")

    data = load_data()

# Standardize coupon codes before duplicate detection.
    if "coupon_code" in data["orders"].columns:

        data["orders"]["coupon_code"] = (
            data["orders"]["coupon_code"]
            .astype("string")
            .str.strip()
            .str.upper()
        )

        print("Standardized orders.coupon_code")


    # Standardize all tables.
    for table_name in data:

        data[table_name] = standardize_columns(
            data[table_name]
        )

        data[table_name] = clean_text_values(
            data[table_name]
        )

        data[table_name] = remove_exact_duplicates(
            data[table_name],
            table_name
        )

    # Standardize category values.
    data = standardize_categories(data)
    verify_categories(data)

    # Convert dates and numbers.
    data = convert_data_types(data)

    # Remove rows with missing mandatory values.
    data = remove_missing_mandatory_values(data)

    # Resolve duplicate primary keys.
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

    # Validate relationships.
    data = validate_relationships(data)

    # Final primary key validation.
    validate_primary_keys(data)

    # Export six cleaned tables.
    export_cleaned_data(data)

    # Combine six tables.
    combined = create_combined_data(data)

    # Export combined dataset.
    export_combined_data(combined)

    # Save conflicting duplicate records for review.
    export_duplicate_review()

    print("\n========== CLEANING COMPLETED ==========\n")


if __name__ == "__main__":
    main()
