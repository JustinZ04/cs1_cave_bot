from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import date, datetime
import time
import calendar
import re


def parse():


    weekday_dict = {'Monday': ('B', 'Q'), 'Tuesday': ('S', 'AI'), 'Wednesday': ('AK', 'AZ'), 'Thursday': ('BB', 'BP'),
                    'Friday': ('BR', 'CH')}
    wb = load_workbook(filename='office_hours.xlsx', data_only=True, read_only=True)

    ws = wb['Office Hours']  # Will have to change the name of the worksheet each time
    ta_list = []
   # print(ws['J22'].value)

    # Get the current weekday
    day = date.today()
    weekday = calendar.day_name[day.weekday()]
    # noon_flag = False
    # add a return here so we don't keep going
    print(weekday)
    if weekday == 'Saturday' or weekday == 'Sunday':
        print('---No TAs have office hours at this time!')
        return None

    cur_time = time.localtime()
    cur_time = time.strftime("%H%M", cur_time)
    cur_time = int(cur_time)
    cur_time = 1515

   # if cur_time < 900 or cur_time > 2100:
    #    print("No office hours.")
     #   return None

    #  if int(cur_time) >= 1200:
    #  noon_flag = True
    print('Current time is ' + str(cur_time))
    time_range = None
    #  Will probably have to edit min/max row for each spreadsheet

    # Should be 43 but upper bound of range is exclusive
    for j in range(5, 44):
       # print("here")
        cell = str(get_column_letter(1) + str(j))
       # print(cell)
        if str(ws[cell].value) != "None":
            s = str(ws[cell].value)
           # print(s)
            s = s.split(' - ')
            s[0] = s[0].replace(":", "")
            s[1] = s[1].replace(":", "")
            s[1] = s[1].replace(" AM", "")
            s[1] = s[1].replace(" PM", "")

            if cur_time > 1259:  # curtime here
                temp = cur_time - 1200  # curtime
            else:
                temp = cur_time  # curtime
            if abs(temp - int(s[0])) < 30:
                print(ws[cell].row)
                time_range = ws[cell].row
                break

    if time_range is None:
        print('+++No TAs have office hours at this time!')
        return None

    found_ta = False
    #  Will probably have to edit this for each spreadsheet

   # cur_time = 1515
    seen_ta = {}
    stime = datetime.now().time()
    file = open("times.log", "a")
    file.write(str(stime) + "\n")
    file.close()
    for i in range(column_index_from_string((weekday_dict[weekday])[0]),
                   column_index_from_string((weekday_dict[weekday])[1])+1):
        for j in range(5, time_range+1):
            col = get_column_letter(i)
            row = str(j)
            cell = "".join((col, row))


           # print(cell)
            s = ws[cell].value
            if type(s) is not str:
                continue




           # print(s)

            office_hour = re.findall(r'\d?\d:\d{2} - \d?\d:\d{2}', s)

            if office_hour is not None:
                if len(office_hour) > 0:
                    office_hour = office_hour[0].split(' - ')

                    office_hour[0] = int(office_hour[0].replace(":", ""))
                    office_hour[1] = int(office_hour[1].replace(":", ""))
                    if office_hour[0] < 900:  # might need to change for new lower TA hour bound
                        office_hour[0] = office_hour[0] + 1200

                    if office_hour[1] < 900:  # might need to change for new lower TA hour bound
                        office_hour[1] = office_hour[1] + 1200

                    if office_hour[0] <= cur_time < office_hour[1]:  # curtime here
                        ta_list.append(ws[cell].value)
                        found_ta = True

    
    if not found_ta:
        print("No TAs have office hours at this time!")
        return None

   # print(len(ta_list))
   # print(str(ta_list))

    etime = datetime.now().time()
    file = open("times.log", "a")
    file.write(str(etime) + "\n")
    file.close()


    return ta_list

if __name__ == '__main__':
    parse()
