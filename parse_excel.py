from openpyxl import load_workbook
from openpyxl.utils import coordinate_from_string, column_index_from_string
from datetime import date
import time
import calendar



def parse():
    weekday_dict = {'Monday': ('B', 'Q'), 'Tuesday': ('S', 'AI'), 'Wednesday': ('AK', 'AZ'), 'Thursday': ('BB', 'BP'),
                    'Friday': ('BR', 'CH')}
    temp = weekday_dict['Monday']
    temp3 = column_index_from_string(temp[0])
    print(temp3)
    time_list = []
    str_time_list = []
    wb = load_workbook(filename='office_hours.xlsx', data_only=True)
    print(wb.sheetnames)
    ws = wb['Office Hours (Deprecated)']
    # print(ws['P15'].value)

    # Get the current weekday
    day = date.today()
    weekday = calendar.day_name[day.weekday()]
    print(weekday)
    cur_time = time.localtime()
    cur_time = time.strftime("%H%M", cur_time)
    if int(cur_time) > 1200:
        cur_time = str(int(cur_time) - 1200)

    #  Will probably have to edit min/max row for each spreadsheet
    for col in ws.iter_cols(min_col=0, max_col=1, min_row=5, max_row=43):
        for cell in col:
            if str(cell.value) != "None":
                s = str(cell.value)
                print(s)
                s = s.split(' - ')
                s[0] = s[0].replace(":", "")
                s[1] = s[1].replace(":", "")
                s[1] = s[1].replace(" AM", "")
                s[1] = s[1].replace(" PM", "")

                if abs(435 - int(s[0])) < 30:
                    print(cell.row)
                    break

    print(ws['BV25'].internal_value)
    #  Will probably have to edit this for each spreadsheet


'''    for i in range(25):
        s = str(cell[i].value)
        # print(s)

        if s != "None":
            #  print(s)
            s = s.split(' - ')
            s[0] = s[0].replace(":", "")
            s[1] = s[1].replace(":", "")
            s[1] = s[1].replace(" AM", "")
            s[1] = s[1].replace(" PM", "")
            time_list.append(s)
    print(time_list)

    if int(cur_time) > 1200:
        cur_time = str(int(cur_time) - 1200)

    not_found = True
    i = 0
    while not_found:
        if abs(355 - int(time_list[i][0])) < 30:
            not_found = False
            print(time_list[i])
            #  Will probably have to edit these for each spreadsheet
            if 15 <= i < 20:
                correct_row = i + 4
                print("entering first if")
            elif 20 <= i:
                print("entering second if")
                correct_row = i + 5
            else:
                print("entering third if")
                correct_row = i + 3

        else:
            i = i + 1
    print(correct_row)
'''
#    print(type(cur_time))
#

# Search through the first few cells to find the correct column
#  Will probably have to edit this for each spreadsheet
'''
    for cell in ws.iter_cols(min_col=0, max_col=31, min_row=0, max_row=2):
        #  print(cell[1].value)
        #  print(cell)
        if cell[1].value == weekday:
            print("cell 1" + str(cell[1]))

    s[0] = s[0].replace(":", "")
    s[1] = s[1].replace(":", "")
    s[1] = s[1].replace(" AM", "")
    s[1] = s[1].replace(" PM", "")
    print(s)
'''

if __name__ == '__main__':
    parse()
