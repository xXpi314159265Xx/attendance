import openpyxl

def main():
    wb = openpyxl.load_workbook('attendance.xlsx')
    sheet = wb.get_sheet_by_name('attendance')
    #print(sheet['A1'].value)
    #print("Cell {} is {}".format(sheet['A1'].coordinate, sheet['A1'].value))

    #a = sheet.max_row
    #print(a)

    attendance_dict = absences(sheet, course_list(sheet))
    print_attendance(sheet, attendance_dict)

def course_list(sheet):
    '''Return a dictionary:
    key: course names
    values: list of 0 absences and tardies'''
    courses = {}
    items = set()
    for cellObj in sheet['C']:
        if cellObj.value == 'Course Code' or cellObj.value == None:
            continue
        else:
            items.add(cellObj.value)
    for name in items:
        courses[name] = [0,0]
    return courses


def absences(sheet, course_dict):
    '''Fills course dictionary with absences and tardies.'''
    for value in sheet['L']:
        if value.value == 'absent' or value.value == 'late':
            row_number = value.row
            cell = 'C'+str(row_number)
            course = sheet[cell].value
            if course == None:
                continue
            elif value.value == 'absent':
                course_dict[course][0] += 1
            else:
                course_dict[course][1] += 1
    return course_dict


def print_attendance(sheet, course_dict):
    '''Prints student name with number of absences and tardies per course.'''
    print()
    print("Attendance for {}".format(sheet['J2'].value))
    print()
    for key, value in course_dict.items():
        print("Absent {} times and late {} times in {}".format(value[0], value[1], key))


main()