from typing import List, Union

from openpyxl import load_workbook
from datetime import date
import calendar


def parse():
    wb = load_workbook(filename='office_hours.xlsx')
    ws = wb.active

    # Get the current weekday
    day = date.today()
    weekday = calendar.day_name[day.weekday()]
    print(weekday)

    # Search through the first few cells to find the correct column
    for cell in ws.iter_cols(min_col=0, max_col=31, min_row=0, max_row=2):
        #  print(cell[1].value)
        if cell[1].value == weekday:
            print(cell[1].value)


if __name__ == '__main__':
    parse()
