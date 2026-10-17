
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 1. Load the raw Excel dataset
input_file = "sales_data.xlsx"
df = pd.read_excel(input_file)

print("DATA CLEANING AND REPORTING AUTOMATION")
print("-" * 45)
print("Original number of rows:", len(df))

# 2. Capture data quality before cleaning
original_rows = len(df)
missing_before = int(df.isnull().sum().sum())
duplicates_before = int(df.duplicated().sum())

print("Missing values before cleaning:", missing_before)
print("Duplicate rows before cleaning:", duplicates_before)

# 3. Remove duplicate rows
df = df.drop_duplicates().copy()

# 4. Standardize text columns
for column in ["Product", "Region", "Customer_ID"]:
    df[column] = df[column].astype("string").str.strip()
    df[column] = df[column].replace("", pd.NA)

df["Product"] = df["Product"].str.title()
df["Region"] = df["Region"].str.title()

# 5. Convert columns to correct data types
df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")

numeric_columns = ["Quantity", "Unit_Price", "Sales"]
for column in numeric_columns:
    df[column] = pd.to_numeric(df[column], errors="coerce")

# 6. Replace invalid values with missing values
df.loc[df["Quantity"] <= 0, "Quantity"] = np.nan
df.loc[df["Unit_Price"] <= 0, "Unit_Price"] = np.nan
df.loc[df["Sales"] < 0, "Sales"] = np.nan

# 7. Fill missing values
df["Customer_ID"] = df["Customer_ID"].fillna("Unknown")
df["Region"] = df["Region"].fillna("Unknown")
df["Product"] = df["Product"].fillna("Unknown")

for column in ["Quantity", "Unit_Price"]:
    df[column] = df[column].fillna(df[column].median())

# Recalculate Sales consistently from quantity and unit price
df["Sales"] = df["Quantity"] * df["Unit_Price"]

# Remove rows where the order date is invalid
df = df.dropna(subset=["Order_Date"]).copy()

df["Quantity"] = df["Quantity"].round().astype(int)
df["Sales"] = df["Sales"].round(2)

# 8. Create a data quality report
missing_after = int(df.isnull().sum().sum())
duplicates_after = int(df.duplicated().sum())

quality_report = pd.DataFrame({
    "Metric": [
        "Original rows",
        "Rows after cleaning",
        "Duplicate rows before cleaning",
        "Duplicate rows after cleaning",
        "Missing cells before cleaning",
        "Missing cells after cleaning",
        "Rows removed"
    ],
    "Value": [
        original_rows,
        len(df),
        duplicates_before,
        duplicates_after,
        missing_before,
        missing_after,
        original_rows - len(df)
    ]
})

# 9. Create an automated sales summary
sales_by_product = (
    df.groupby("Product", as_index=False)
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Quantity=("Quantity", "sum"),
        Number_of_Orders=("Order_ID", "count")
    )
    .sort_values("Total_Sales", ascending=False)
)

sales_by_region = (
    df.groupby("Region", as_index=False)
    .agg(Total_Sales=("Sales", "sum"))
    .sort_values("Total_Sales", ascending=False)
)

# 10. Export cleaned data and reports to Excel
with pd.ExcelWriter("cleaned_sales_data.xlsx") as writer:
    df.to_excel(writer, sheet_name="Cleaned_Data", index=False)

with pd.ExcelWriter("data_quality_report.xlsx") as writer:
    quality_report.to_excel(writer, sheet_name="Quality_Report", index=False)
    sales_by_product.to_excel(writer, sheet_name="Sales_By_Product", index=False)
    sales_by_region.to_excel(writer, sheet_name="Sales_By_Region", index=False)

# 11. Generate a visual summary chart
plt.figure(figsize=(9, 5))
plt.bar(sales_by_product["Product"], sales_by_product["Total_Sales"])
plt.title("Total Sales by Product After Cleaning")
plt.xlabel("Product")
plt.ylabel("Total Sales")
plt.xticks(rotation=25)
plt.tight_layout()
plt.savefig("cleaning_summary.png", dpi=300)
plt.close()

# 12. Print the final results
print("\nCLEANING COMPLETED SUCCESSFULLY!")
print("Rows after cleaning:", len(df))
print("Duplicate rows remaining:", duplicates_after)
print("\nSales by product:")
print(sales_by_product.to_string(index=False))
print("\nFiles created:")
print("1. cleaned_sales_data.xlsx")
print("2. data_quality_report.xlsx")
print("3. cleaning_summary.png")
