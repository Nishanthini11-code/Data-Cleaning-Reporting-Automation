# Data Cleaning and Reporting Automation

## Project Overview
This project automates data cleaning and reporting using Python. It processes sales data, handles missing values and duplicate records, standardizes inconsistent data, and generates Excel reports and a visual summary.

## Objectives
- Identify and handle missing values.
- Remove duplicate records.
- Standardize inconsistent data.
- Clean and validate sales data.
- Generate automated Excel reports.
- Visualize data-cleaning results.

## Technologies Used
- Python
- Pandas
- NumPy
- Matplotlib
- OpenPyXL
- Microsoft Excel

## Project Files
- `data_cleaning.py` — Python script for data cleaning and reporting.
- `sales_data.xlsx` — Original sales dataset.
- `cleaned_sales_data.xlsx` — Cleaned sales dataset.
- `data_quality_report.xlsx` — Data quality and sales analysis reports.
- `cleaning_summary.png` — Visual summary of the cleaning results.

## Results
- Original records: 262
- Duplicate rows identified: 12
- Records after cleaning: 250
- Duplicate rows remaining: 0
- Missing values identified before cleaning: 26

## How to Run
1. Install Python.
2. Install the required libraries:

   `python -m pip install pandas numpy matplotlib openpyxl`

3. Keep `data_cleaning.py` and `sales_data.xlsx` in the same folder.
4. Run:

   `python data_cleaning.py`

## Conclusion
This project demonstrates how Python can automate data cleaning, improve data quality, and generate reports for sales analysis.
