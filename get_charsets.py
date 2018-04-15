import xlrd
sheet = xlrd.open_workbook('charset.xlsx').sheet_by_index(0)
rows = sheet.nrows
list_ = []
for x in range(0,rows):
    got_value = sheet.cell(x,0).value.strip()
    list_.append(got_value)
print(list_)