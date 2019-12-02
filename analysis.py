'''
Final Project
Sydney Fowler and Matt Hileman
15-12-2019
Description: Allows the user to select from a set of analysis options for each column in their file. For each sheet
selected, a new analysis sheet is created with the results from each analyzed column.
'''

import openpyxl
from cleanup import get_user_selection
from cleanup import print_menu
from cleanup import get_wb_path
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

ANALYSIS_OPTIONS_LIST = ["Sum", "Count", "Max", "Min", "Check Unique", "Average", "No Analysis", "Finish Sheet"]
NO_ANALYSIS = ANALYSIS_OPTIONS_LIST.index("No Analysis")
BREAK_SHEET = ANALYSIS_OPTIONS_LIST.index("Finish Sheet")


def init():
    wb_path = get_wb_path()
    wb = openpyxl.load_workbook(wb_path)
    sheets = wb.get_sheet_names()

    # Initialize 2D dictionary representing each sheet and its column headers
    sheet_header_lookup = {}
    for sheet_name in sheets:
        sheet_header_lookup.setdefault(sheet_name, {})
        sheet = wb.get_sheet_by_name(sheet_name)
        for cell in sheet[1]:
            sheet_header_lookup[sheet_name].setdefault(cell.value, None)

    # Get user selections
    for sheet_name in sheets:
        # Check if user wants to process this sheet
        analyze_sheet = input("Would you like to analyze sheet " + sheet_name + "? (y/n) ")
        if analyze_sheet not in ("yes", "Yes", "Y", "y"):
            continue

        # Get user selections for each header in sheet
        for header in sheet_header_lookup[sheet_name]:
            print_menu(sheet_name, header, ANALYSIS_OPTIONS_LIST)
            user_selection = get_user_selection(ANALYSIS_OPTIONS_LIST)
            if user_selection == BREAK_SHEET:       # Check if user wants to break out of sheet
                break
            elif user_selection == NO_ANALYSIS:     # Check if user wants to skip this column
                continue
            else:
                sheet_header_lookup[sheet_name][header] = user_selection

    # Analyze data
    for sheet_name in sheets:
        sheet = wb.get_sheet_by_name(sheet_name)
        for col in range(1, sheet.max_column + 1):
            analysis_number = sheet_header_lookup[sheet_name][sheet.cell(row=1, column=col).value]
            if analysis_number is not None:
                col_letter = get_column_letter(col)
                perform_analysis(wb, sheet_name, sheet[col_letter], int(analysis_number))

    # Save to a new copy of the workbook
    new_file = wb_path[:len(wb_path) - 5] + "_EDITED.xlsx"
    wb.save(new_file)


def perform_analysis(wb, sheet_name, wb_range, analysis_number):
    if analysis_number == ANALYSIS_OPTIONS_LIST.index("Sum"):
        op = "SUM"
        ret_val = perform_sum(sheet_name, wb_range)
    elif analysis_number == ANALYSIS_OPTIONS_LIST.index("Count"):
        op = "COUNT"
        ret_val = perform_count(sheet_name, wb_range)
    elif analysis_number == ANALYSIS_OPTIONS_LIST.index("Max"):
        op = "MAX"
        ret_val = perform_max(sheet_name, wb_range)
    elif analysis_number == ANALYSIS_OPTIONS_LIST.index("Min"):
        op = "MIN"
        ret_val = perform_min(sheet_name, wb_range)
    elif analysis_number == ANALYSIS_OPTIONS_LIST.index("Check Unique"):
        op = "UNIQUE"
        ret_val = perform_unique(sheet_name, wb_range)
    elif analysis_number == ANALYSIS_OPTIONS_LIST.index("Average"):
        op = "AVERAGE"
        ret_val = perform_average(sheet_name, wb_range)

    # If analysis sheet is not created, create it
    if sheet_name.upper() + " ANALYSIS" not in wb.get_sheet_names():
        wb.create_sheet(title=sheet_name.upper() + " ANALYSIS")
        sheet = wb.get_sheet_by_name(sheet_name.upper() + " ANALYSIS")
        my_font = Font(bold=True)
        sheet['A1'].font = my_font
        sheet['A1'].value = "ANALYSIS SUMMARY"

    sheet = wb.get_sheet_by_name(sheet_name.upper() + " ANALYSIS")
    new_row = sheet.max_row + 1
    sheet.cell(row=new_row, column=1).value = wb_range[0].value
    sheet.cell(row=new_row, column=2).value = op
    sheet.cell(row=new_row, column=3).value = ret_val


def perform_sum(sheet_name, wb_range):
    ret_sum = 0
    non_float_flag = False

    # Analyze column
    for cell in wb_range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Perform op
        try:
            ret_sum += float(cell.value)
        except Exception:
            non_float_flag = True

    # Print warning message if some of the cell values could not be summed
    if non_float_flag:
        print("WARNING: Some cells in " + sheet_name + ":" + wb_range[0].value + " were not numerical values.")

    return str(ret_sum)


def perform_count(sheet_name, wb_range):
    ret_count = 0

    # Analyze column
    for cell in wb_range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Perform op
        if cell.value:
            ret_count += 1

    return str(ret_count)


def perform_max(sheet_name, wb_range):
    ret_max = 0
    non_float_flag = False

    # Analyze column
    for cell in wb_range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Perform op
        try:
            num = float(cell.value)
            if num > ret_max:
                ret_max = num
        except Exception:
            non_float_flag = True

    # Print warning message if some of the cell values could not be summed
    if non_float_flag:
        print("WARNING: Some cells in " + sheet_name + ":" + wb_range[0].value + " were not numerical values.")

    return str(ret_max)


def perform_min(sheet_name, wb_range):
    ret_min = "Uninitialized"
    non_float_flag = False

    # Analyze column
    for cell in wb_range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Perform op
        try:
            num = float(cell.value)
            if ret_min == "Uninitialized":
                ret_min = num
            elif num < ret_min:
                ret_min = num
        except Exception:
            non_float_flag = True

    # Print warning message if some of the cell values could not be summed
    if non_float_flag:
        print("WARNING: Some cells in " + sheet_name + ":" + wb_range[0].value + " were not numerical values.")

    return str(ret_min)


def perform_unique(sheet_name, wb_range):
    entries = []
    # Analyze column
    for cell in wb_range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Perform op
        if cell.value:
            if cell.value.lower() in entries:
                return "False"
            entries.append(cell.value.lower())

    return "True"


def perform_average(sheet_name, wb_range):
    init_sum = 0
    init_count = 0
    non_float_flag = False

    # Analyze column
    for cell in wb_range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Perform op
        try:
            init_sum += float(cell.value)
            init_count += 1
        except Exception:
            non_float_flag = True

    # Print warning message if some of the cell values could not be summed
    if non_float_flag:
        print("WARNING: Some cells in " + sheet_name + ":" + wb_range[0].value + " were not numerical values.")

    if init_count == 0:
        return 0
    else:
        return str(init_sum / init_count)


init()
