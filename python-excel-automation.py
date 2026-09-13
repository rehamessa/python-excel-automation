#-------------------------------------------------------
#open the spreedsheets
#-------------------------------------------------------

from openpyxl import load_workbook
workbook=load_workbook('sales.xlsx')
sheet=workbook.active

#-------------------------------------------------------
#Read all sales value
#-------------------------------------------------------
highest_sales=0
total_sales=0
for row in range(2,sheet.max_row+1):
    sales=sheet[f"B{row}" ] .value

    total_sales+=sales# total sales

    if sales>highest_sales:
        highest_sales=sales #highest sales

average_sale=total_sales/(sheet.max_row-1) #average sales

product_count=sheet.max_row-1 #calculate product count




    