from openpyxl import load_workbook

wb = load_workbook('sagatave_eksamenam.xlsx', data_only=True)
ws = wb['Page_0']  # Corrected sheet name

max_row = ws.max_row
count = 0

for row in range(2, max_row + 1):
    priority = ws['H' + str(row)].value
    delivery_date = ws['J' + str(row)].value

    if priority == 'High' and hasattr(delivery_date, 'year') and delivery_date.year == 2015:
        count += 1

print(count)
