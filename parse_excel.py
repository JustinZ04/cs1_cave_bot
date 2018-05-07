from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from datetime import date, datetime
import time
import calendar
import re


def parse():
    stime = datetime.now().time()
    file = open("times.log", "a")
    file.write(str(stime) + "\n")
    file.close()

    weekday_dict = {'Monday': ('B', 'Q'), 'Tuesday': ('S', 'AI'), 'Wednesday': ('AK', 'AZ'), 'Thursday': ('BB', 'BP'),
                    'Friday': ('BR', 'CH')}
    wb = load_workbook(filename='office_hours.xlsx', data_only=True)
    etime = datetime.now().time()
    file = open("times.log", "a")
    file.write(str(etime) + "\n")
    file.close()
    ws = wb['Office Hours']  # Will have to change the name of the worksheet each time
    ta_list = []
    # print(ws['P15'].value)

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

    if cur_time < 900 or cur_time > 2100:
        print("No office hours.")
        return None

    #  if int(cur_time) >= 1200:
    #  noon_flag = True
    print('Current time is ' + str(cur_time))
    time_range = None
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

                if cur_time > 1259:  # curtime here
                    temp = cur_time - 1200  # curtime
                else:
                    temp = cur_time  # curtime
                if abs(temp - int(s[0])) < 30:
                    print(cell.row)
                    time_range = cell.row
                    break
    if time_range is None:
        print('+++No TAs have office hours at this time!')
        return None



    found_ta = False
    #  Will probably have to edit this for each spreadsheet



    for col in ws.iter_cols(min_col=column_index_from_string((weekday_dict[weekday])[0]),  # weekday here
                            max_col=column_index_from_string((weekday_dict[weekday])[1]), min_row=5,
                            max_row=time_range):
        for cell in col:
            if str(cell.value) != "None":
                s = str(cell.value)
                # print(s)
                office_hour = re.findall(r'\d?\d:\d{2} - \d?\d:\d{2}', s)
                if re.findall(r'Lecture', s):
                    office_hour = None
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
                            ta_list.append(cell.value)
                            found_ta = True



    if not found_ta:
        print("No TAs have office hours at this time!")
        return None



    return ta_list


if __name__ == '__main__':
    parse()
