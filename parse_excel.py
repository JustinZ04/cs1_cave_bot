from openpyxl import load_workbook
from datetime import date
import calendar


def parse():
    wb = load_workbook(filename='office_hours.xlsx', data_only=True)
    ws = wb.active

    # Get the current weekday
    day = date.today()
    weekday = calendar.day_name[day.weekday()]
    print(weekday)

    # Search through the first few cells to find the correct column
    for cell in ws.iter_cols(min_col=0, max_col=31, min_row=0, max_row=2):
        #  print(cell[1].value)
        print(cell)
        if cell[1].value == weekday:
            print(cell[1].value)

    #  This wont work
    if ws['A3'].value == '8:30':
        print('matched')

    #  I think this is a better way
    s = str(ws['A3'].value)
    s = s[:4]
    s = s.replace(":", "")
    print(s)

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
