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

#-------------------------------------------------------
#create new excel report
#-------------------------------------------------------

from openpyxl import Workbook

report_workbook = Workbook()

report_sheet = report_workbook.active

#-------------------------------------------------------
#wrirting report data
#-------------------------------------------------------

#title

report_sheet['A1']='Metrics'
report_sheet['B1']='values'

#report values

report_sheet['A2']='total sales'
report_sheet['B2']=total_sales

report_sheet['A3']='highest sales'
report_sheet['B3']=highest_sales

report_sheet['A4']='average sales'
report_sheet['B4']=average_sale

report_sheet['A4']='product count'
report_sheet['B4']=product_count







    