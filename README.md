# python-excel-automation
# Excel Sales Report Generator

A Python script that reads raw sales data from an Excel file, calculates key metrics (total, average, highest sale, product count), and generates a new summary report as a separate Excel file — using `openpyxl`.


 Creates a new Excel workbook (`sales_report.xlsx`) and writes the calculated metrics into it

## Example Output

Given the input data above, the generated `sales_report.xlsx` contains:

| Metric          | Value |
|-----------------|-------|
| Total Sales     | 2900  |
| Average Sales   | 725   |
| Highest Sale    | 1200  |
| Products Count  | 4     |


## Python Concepts Used

- Reading Excel files with `load_workbook()`
- Creating new Excel files with `Workbook()`
- Looping through rows with `sheet.max_row`
- Aggregation logic (sum, average, max, count)
- Writing calculated values to specific cells
