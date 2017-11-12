import openpyxl

def main():
    wb = openpyxl.load_workbook('attendance.xlsx')
    sheet = wb.get_sheet_by_name('attendance')
    #print(sheet['A1'].value)
    #print("Cell {} is {}".format(sheet['A1'].coordinate, sheet['A1'].value))

    a = sheet.max_row
    #print(a)

    attendance_dict = absences(sheet, course_list(sheet))
    print_attendance(sheet, attendance_dict)

def course_list(sheet):
    courses = {}
    items = set()
    for cellObj in sheet['C']:
        if cellObj.value == 'Course Code':# or cellObj.value == None:
            continue
        else:
            items.add(cellObj.value)
    for name in items:
        courses[name] = 0
    #print(items)
    #print(courses)
    return courses


def absences(sheet, course_dict):
    #Beginning to tally absenses
    #Check for absent value or tardy value
    #Find what row it's in
    #Go to column C of that row and find that value
    #Update the dictionary value of that value
    for value in sheet['L']:
        if value.value == 'absent':
            row_number = value.row
            cell = 'C'+str(row_number)
            course = sheet[cell].value
            #print(course)
            course_dict[course] += 1
    #print(course_dict)
    return course_dict


def print_attendance(sheet, course_dict):
    print()
    print("Attendance for {}".format(sheet['J2'].value))
    print()
    for key, value in course_dict.items():
        print("Absent {} times in {}".format(value, key))


main()