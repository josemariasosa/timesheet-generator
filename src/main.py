#!/usr/bin/env/ python3

from tabulate import tabulate
from datetime import date, timedelta
from openpyxl import load_workbook

import os
import argparse

parser = argparse.ArgumentParser(description='List of hours [0-8], at least 5 days.')
parser.add_argument('hours', type=int, help='Consecutive week hours Eg. 88754, 5510030')
args = parser.parse_args()

files = sorted(os.listdir('files/'))

DATEFORMAT = '%b-%d-%y'

CELLMAPS = [
    {'data': 'date', 'day': 0, 'cells': ['B12', 'A18']},
    {'data': 'hour', 'day': 0, 'cells': ['A20']},
    {'data': 'date', 'day': 1, 'cells': ['B18']},
    {'data': 'hour', 'day': 1, 'cells': ['B20']},
    {'data': 'date', 'day': 2, 'cells': ['C18']},
    {'data': 'hour', 'day': 2, 'cells': ['C20']},
    {'data': 'date', 'day': 3, 'cells': ['D18']},
    {'data': 'hour', 'day': 3, 'cells': ['D20']},
    {'data': 'date', 'day': 4, 'cells': ['E18']},
    {'data': 'hour', 'day': 4, 'cells': ['E20']},
    {'data': 'date', 'day': 5, 'cells': ['F18']},
    {'data': 'hour', 'day': 5, 'cells': ['F20']},
    {'data': 'date', 'day': 6, 'cells': ['G18']},
    {'data': 'hour', 'day': 6, 'cells': ['G20']},
    {'data': 'today', 'cells': ['G26']},
    {'data': 'total', 'cells': ['G22']},
]

def main():
    hours = format_hours(list(str(args.hours)))
    today = date.today()
    monday = get_current_monday(today)
    last_week_monday = monday - timedelta(days=7)
    formatted_dates = [
        format_date(last_week_monday + timedelta(days=x))
        for x in range(7)
    ]
    confirm_user_input(hours, formatted_dates)
    
    base = get_base_file_name()
    output_file_name = get_output_file_name(monday, base)

    workbook = load_workbook(filename=f'files/{base}')
    sheet = workbook.active

    for cellmap in CELLMAPS:
        data = cellmap.get('data')
        for cell in cellmap.get('cells'):
            if data == 'date':
                sheet[cell] = formatted_dates[cellmap.get('day')]
            elif data == 'hour':
                sheet[cell] = hours[cellmap.get('day')]
            elif data == 'total':
                sheet[cell] = sum(hours)
            elif data == 'today':
                sheet[cell] = today.strftime(DATEFORMAT)
            else:
                raise ValueError(f'Incorrect CELLMAPS data {data}')

    workbook.save(filename=f'files/{output_file_name}')
    print(f'New file was created: {output_file_name}')

def get_output_file_name(monday, base) -> str:
    tail = monday.strftime('%Y-%m-%d')
    if tail in [f.split('.', 1)[0][-10:] for f in files]:
        print(f'\nFile for monday {tail} already exists!')
        if not input('Replace? (y/n): ').strip().lower().startswith('y'):
            print('bye.')
            exit()
    next_number = '{:02d}'.format(max([int(f[:2]) for f in files if f[:2].isdigit()]) + 1)
    file_name = base.split('.', 1)[0][2:-10]
    file_extension = base.split('.', 1)[1] 
    return next_number + file_name + tail + '.' + file_extension

def get_base_file_name():
    base = [f for f in files if f.startswith('00-')]
    assert len(base) == 1, 'Only 1 template, starting with 00-'
    return base[0]

def format_hours(hours):
    assert len(hours) <= 7, 'Max days 7: {}'.format(hours)
    result = [int(hour) for hour in hours] + [0 for x in range(7 - len(hours))]
    assert all([0 <= res <= 8 for res in result]), 'Max hours per day 8!'
    return result

def confirm_user_input(hours, dates):
    print(tabulate({"Date": dates, "Hours": hours}, headers="keys"))
    print('---------  -------')
    print(f'   Total:   {sum(hours)} hrs')
    if input('\nCorrect? (y/n): ').strip().lower().startswith('y'):
        return True
    print('try again.')
    exit()

def format_date(d):
    return d.strftime(DATEFORMAT)

def get_current_monday(today):
    return today - timedelta(days=today.weekday())

if __name__=='__main__':
    main()
    exit()
 