import openpyxl

book = openpyxl.load_workbook('C:/Users/Subodh Maharjan/Desktop/fus/Rec.xlsx')

sheet = book.active

a1 = sheet['A1']

a3 = sheet.cell(row=2, column=6)

print(a1.value)

print(a3.value)