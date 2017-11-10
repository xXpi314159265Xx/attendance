import openpyxl

wb = openpyxl.load_workbook('attendance.xlsx')
sheet = wb.get_sheet_by_name('attendance')
print(sheet['A1'].value)
print("Cell {} is {}".format(sheet['A1'].coordinate, sheet['A1'].value))

a = sheet.max_row
print(a)

items = set()
for cellObj in sheet['C']:
    if cellObj.value == 'Course Code' or cellObj.value == None:
        continue
    else:
        items.add(cellObj.value)
print(items)
courses = {}
for name in items:
    courses[name] = 0
print(courses)

#Beginning to tally absenses
#Check for absent value or tardy value
#Find what row it's in
#Go to column C of that row and find that value
#Update the dictionary value of that value
for value in sheet['L']:
    if value == 'absent':

courses['Discrete Mathematics'] += 1
print(courses)