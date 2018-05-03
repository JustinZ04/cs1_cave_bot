from openpyxl import load_workbook
from datetime import date
import time
import calendar


def parse():
    time_list = []
    str_time_list = []
    wb = load_workbook(filename='office_hours.xlsx', data_only=True)
    ws = wb.active

    # Get the current weekday
    day = date.today()
    cur_time = time.localtime()
    cur_time = time.strftime("%H%M", cur_time)

    for cell in ws.iter_cols(min_col=0, max_col=1, min_row=3, max_row=27):
        continue

    print(cell)
    for i in range(25):
        s = str(cell[i].value)
        # print(s)

        if s != "None":
            print(s)
            s = s.split(' - ')
        #  print(len(s))
            s[0] = s[0].replace(":", "")
        #  print(s[1])
            s[1] = s[1].replace(":", "")
            s[1] = s[1].replace(" AM", "")
            s[1] = s[1].replace(" PM", "")
            time_list.append(s)
    print(time_list)

    if int(cur_time) > 1200:
        cur_time = str(int(cur_time) - 1200)

    print(type(cur_time))
    weekday = calendar.day_name[day.weekday()]
    print(weekday)

    # Search through the first few cells to find the correct column
    for cell in ws.iter_cols(min_col=0, max_col=31, min_row=0, max_row=2):
        #  print(cell[1].value)
        #  print(cell)
        if cell[1].value == weekday:
            print(cell[1].value)

    #  I think this is a better way
    s = str(ws['A3'].value)
    s = s.split(' - ')
    print(s)

    s[0] = s[0].replace(":", "")
    s[1] = s[1].replace(":", "")
    s[1] = s[1].replace(" AM", "")
    s[1] = s[1].replace(" PM", "")
    print(s)

    #  s = s[:4]
    #  s = s.replace(":", "")
    #  print(s)

    #  I really hope you know a better way to do this or
    #  no one is ever going to be able to see this code
    #  s = s.replace(":", "")
    #  s = s.replace(" - ", "")
    #  s = s.replace("A", "")
    #  s = s.replace("P", "")
    #  s = s.replace("M", "")
    #  s = s[3:]
    #  print(s)


if __name__ == '__main__':
    parse()
