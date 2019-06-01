#!/usr/bin/env Python3
import PySimpleGUI
import re
import xlsxwriter
import datetime
from pathlib import Path
import xlrd
from xlutils.copy import copy as xl_copy


def nearest(compare_dates, pivot):
    return min(compare_dates, key=lambda k: abs(k - pivot))


def duplicate(input_items):
    unique = []
    for item_loop in input_items:
        if item_loop not in unique:
            unique.append(item_loop)
    return unique


def get_user_date(user_input_values):
    user_start_year = user_input_values[0][user_input_values[0].rfind("/") + 1:]
    user_start_month = user_input_values[0][user_input_values[0].find("/") + 1:user_input_values[0].rfind("/")]
    user_start_day = user_input_values[0][:user_input_values[0].find("/")]
    user_start_date = datetime.datetime(int(user_start_year), int(user_start_month), int(user_start_day))

    user_end_year = user_input_values[1][user_input_values[1].rfind("/") + 1:]
    user_end_month = user_input_values[1][user_input_values[1].find("/") + 1:user_input_values[1].rfind("/")]
    user_end_day = user_input_values[1][:user_input_values[1].find("/")]
    user_end_date = datetime.datetime(int(user_end_year), int(user_end_month), int(user_end_day))
    return user_start_date, user_end_date


def setup_gui(gui_control):
    gui_control.ChangeLookAndFeel('TealMono')

    layout = [
        [gui_control.Text('Start Date: ', size=(17, 1), justification='left'),
         gui_control.Text('End Date: ', size=(20, 1), justification='left')],
        [gui_control.InputText('23/09/2017', size=(20, 1)), gui_control.InputText('23/09/2017', size=(20, 1))],
        [gui_control.Text('EVODiag Log File: ', size=(20, 1), auto_size_text=False, justification='left')],
        [gui_control.InputText('SCFVUJAW7LPX93121.log', justification='left'), gui_control.FileBrowse()],
        [gui_control.Submit(tooltip='Click to submit this window'), gui_control.Cancel()]
    ]

    window = gui_control.Window('DTC Log Parser', layout, default_element_size=(40, 1), grab_anywhere=False)

    user_value = window.Read()[1]
    print(user_value)
    return user_value


Hex_Values = []
Hex_Count = []
VIN = []
Dates_Found = []
Dates_DDMMYYYY = []
Start_Line = []

values = setup_gui(PySimpleGUI)
print(values[0])

if values[0][3] == "0":
    values[0] = values[0][0:3] + values[0][4:]

VIN = values[2][values[2].rfind("/") + 1:-4]

Log_File = open(values[2], 'r').read()
Log_File = Log_File.splitlines()

for (index, Line) in enumerate(Log_File):
    x = re.search(r'\d+/\d+/\d+', Line)
    if x:
        Dates_Found.append(x.group())

(User_Start_Date, User_End_Date) = get_user_date(values)

for dates in Dates_Found:
    Log_Year = dates[dates.rfind("/") + 1:]
    Log_Month = dates[dates.find("/") + 1:dates.rfind("/")]
    Log_Day = dates[:dates.find("/")]
    Log_Date = datetime.datetime(int(Log_Year), int(Log_Month), int(Log_Day))
    if Log_Date >= User_Start_Date:
        Dates_DDMMYYYY.append(Log_Date)

Date_To_Read_From = nearest(Dates_DDMMYYYY, User_Start_Date)
Date_To_Read_From = str(Date_To_Read_From.strftime("%d/%m/%Y"))

if Date_To_Read_From[3] == "0":
    Date_To_Read_From = Date_To_Read_From[0:3] + Date_To_Read_From[4:]

for (index, Line) in enumerate(Log_File):
    x = re.search(r'\d+/\d+/\d+', Line)
    if x:
        if Date_To_Read_From in Line:
            Start_Line = index
            break

Log_File = Log_File[Start_Line:]

for (idx, Line) in enumerate(Log_File):
    x = re.match(r"[0-9A-F]+\s[PCBU][0-9A-F]+", Line)
    if x:
        # if int(x.span()[1]) - int(x.span()[0]) == 8:
        DTC_String = x.group()
        Space_Index = DTC_String.find(' ')
        Hex_Values.append(DTC_String[Space_Index + 1:])

Unique_Hex = duplicate(Hex_Values)

for idx, items in enumerate(Unique_Hex):
    Hex_Count.append(Hex_Values.count(items))

Print_Values = zip(Unique_Hex, Hex_Count)

config = Path(values[2][0:values[2].rfind("/") + 1] + VIN + ' Parsed DTCs.xls')

today = datetime.datetime.today()
d1 = today.strftime("%d-%m-%Y")

if not config.is_file():
    workbook = xlsxwriter.Workbook(values[2][0:values[2].rfind("/") + 1] + VIN + ' Parsed DTCs.xls')
    workbook.close()

rb = xlrd.open_workbook((values[2][0:values[2].rfind("/") + 1] + VIN + ' Parsed DTCs.xls'))
wb = xl_copy(rb)

if d1 not in rb.sheet_names():
    if "Sheet1" in rb.sheet_names():
        idx = rb.sheet_names().index('Sheet1')
        wb.get_sheet(idx).name = d1
        Sheet1 = wb.get_sheet(idx)
    else:
        Sheet1 = wb.add_sheet(d1)
else:
    Sheet1 = wb.get_sheet(d1)
    Sheet_Names = rb.sheet_names()
    D1_Index = Sheet_Names.index(d1)

    for row in range(rb.sheet_by_index(D1_Index).nrows):
        for column in range(rb.sheet_by_index(D1_Index).ncols):
            Sheet1.write(row, column, '')

row = 3
Sheet1.write(0, 0, 'From: ')
Sheet1.write(0, 1, values[0])
Sheet1.write(1, 0, 'DTC: ')
Sheet1.write(1, 1, 'No. Occurrences')

for DTCs, Counts in Print_Values:
    Sheet1.write(row, 0, DTCs)
    Sheet1.write(row, 1, Counts)
    row += 1

print(VIN)
wb.save(VIN + ' Parsed DTCs.xls')