from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import date
import time
import calendar
import re


def parse():
    # may need to change tuple bounds in this dictionary to account for the new spreadsheet
    weekday_dict = {'Monday': ('D', 'N'), 'Tuesday': ('T', 'AH'), 'Wednesday': ('AM', 'AW'), 'Thursday': ('BC', 'BO'),
                    'Friday': ('BT', 'CD')}
    wb = load_workbook(filename='office_hours.xlsx', data_only=True, read_only=True)

    ws = wb['Office Hours']  # Will have to change the name of the worksheet each time
    ta_list = []

    # Get the current weekday
    day = date.today()
    weekday = calendar.day_name[day.weekday()]

    if weekday == 'Saturday' or weekday == 'Sunday':
        print('No TAs have office hours at this time!')
        return None

    cur_time = time.localtime()
    cur_time = time.strftime("%H%M", cur_time)
    cur_time = int(cur_time)

    if cur_time < 900 or cur_time > 2100:
        return None

    time_range = None

    # Should be 43 but upper bound of range is exclusive
    for j in range(5, 44):
        cell = str(get_column_letter(1) + str(j))

        if str(ws[cell].value) != "None":
            s = str(ws[cell].value)
            s = s.split(' - ')
            s[0] = s[0].replace(":", "")
            s[1] = s[1].replace(":", "")
            s[1] = s[1].replace(" AM", "")
            s[1] = s[1].replace(" PM", "")

            if cur_time > 1259:
                temp = cur_time - 1200
            else:
                temp = cur_time
            if abs(temp - int(s[0])) < 30:
                #  print(ws[cell].row)
                time_range = ws[cell].row
                break

    if time_range is None:
        print('No TAs have office hours at this time!')
        return None

    found_ta = False

    i = column_index_from_string((weekday_dict[weekday])[0])
    while i <= column_index_from_string((weekday_dict[weekday])[1]):
        j = time_range

        while j >= 5:  # may need to change the 5 to account for lower bound of summer spreadsheet
            col = get_column_letter(i)
            row = str(j)
            cell = "".join((col, row))
            s = ws[cell].value

            if s is None:
                j -= 1
                continue

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

                    if office_hour[0] <= cur_time < office_hour[1]:
                        ta_list.append(ws[cell].value)
                        found_ta = True
                        i += 1
                        break
                    else:
                        i += 1
                        break
            j = j - 1
        i = i + 1

    if not found_ta:
        print("No TAs have office hours at this time!")
        return None

    return ta_list


if __name__ == '__main__':
    parse()
