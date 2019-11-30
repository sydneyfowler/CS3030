'''
Final Project
Sydney Fowler and Matt Hileman
15-12-2019
Description: Allows the user to select a set of cleanup rules for each column in their file and applies said cleanup
to a new version of the file.
'''

import os
import sys
import re
import openpyxl
import custom_dictionaries
import pprint
from openpyxl.utils import get_column_letter
from datetime import datetime


CLEANUP_OPTIONS_LIST = ["Cleanup Phone Numbers", "Cleanup Email Addresses", "Cleanup States", "Cleanup Zip Codes",
                        "Cleanup Dates", "Cleanup Web Address", "Cleanup Social Media",
                        "Produce List of Unique Entries", "Check Entries Against List", "Truncate to Character Limit",
                        "Check Data Type", "No Cleaning", "Finish Sheet"]
DATA_TYPE_LIST = ["Whole Number", "Decimal Value", "Currency", "Text String", "Datetime Stamp", "Not Specified"]

NO_CLEANING = CLEANUP_OPTIONS_LIST.index("No Cleaning")
BREAK_SHEET = CLEANUP_OPTIONS_LIST.index("Finish Sheet")


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
        process_sheet = input("Would you like to clean sheet " + sheet_name + "? (y/n) ")
        if process_sheet not in ("yes", "Yes", "Y", "y"):
            continue

        # Get user selections for each header in sheet
        for header in sheet_header_lookup[sheet_name]:
            print_menu(sheet_name, header)
            user_selection = get_user_selection(CLEANUP_OPTIONS_LIST)
            if user_selection == BREAK_SHEET:       # Check if user wants to break out of sheet
                break
            elif user_selection == NO_CLEANING:     # Check if user wants to skip this column
                continue
            else:
                sheet_header_lookup[sheet_name][header] = user_selection

    # Process data
    for sheet_name in sheets:
        sheet = wb.get_sheet_by_name(sheet_name)
        for col in range(1, sheet.max_column + 1):
            process_number = sheet_header_lookup[sheet_name][sheet.cell(row=1, column=col).value]
            if process_number is not None:
                col_letter = get_column_letter(col)
                process_column(sheet[col_letter],
                               int(process_number))

    # Save to a new copy of the workbook
    new_file = wb_path[:len(wb_path) - 5] + "_EDITED.xlsx"
    wb.save(new_file)


def get_wb_path():
    while (True):  # Loop until you get a valid Excel file
        wb_path = input("Type path of your Excel file: ")
        if os.path.exists(wb_path):
            if wb_path[-5:] != ".xlsx":
                print("ERROR: Must be a .xlsx file.")
            else:
                break
        else:
            print("ERROR: Invalid file path.")

    return wb_path


def print_menu(sheet_name, header):
    print('-' * 40)
    print("SHEET: " + str(sheet_name))
    print("HEADER: " + str(header))
    print("Select an option (0-" + str(len(CLEANUP_OPTIONS_LIST) - 1) + ")")
    print('-' * 40)
    # Prints each item in list
    for item in CLEANUP_OPTIONS_LIST:
        print("(" + str(CLEANUP_OPTIONS_LIST.index(item)) + ") " + item)


def get_user_selection(l):
    # Error checking loop - input is an integer and is a valid menu item
    while (True):
        print('-' * 40)
        print("Choice: ", end='')

        # Initialize choice
        user_choice = input()

        # Error handling: makes sure input is integer, stores interger
        try:
            user_choice = int(user_choice)
        # If input is not an integer, display error, has user try again.
        except ValueError:
            # Error Message
            print()
            print(str(user_choice) + " is not a valid input (NOT AN INT)")
            print("Choose a numeric value from the options above between (0-" + str(len(l) - 1)
                  + ").")
            continue

        # Error handling: makes sure the user's choice is a valid menu option
        if (user_choice < 0) or (user_choice >= len(l)):
            # Error Message
            print()
            print(str(user_choice) + " is not a valid input.")
            print("Choose a numeric value from the options above between (0-" + str(len(l) - 1)
                  + ").")
            continue

        # If input is valid, return the input value, break from error loop
        else:
            return user_choice


def process_column(range, process_number):
    if process_number == 0:
        clean_phone_number(range)
    elif process_number == 1:
        clean_email_address(range)
    elif process_number == 2:
        clean_states(range)
    elif process_number == 3:
        clean_zip_codes(range)
    elif process_number == 4:
        clean_dates(range)
    elif process_number == 5:
        clean_web_addresses(range)
    elif process_number == 6:
        clean_social_media(range)
    elif process_number == 7:
        get_unique_entries(range)
    elif process_number == 8:
        user_list = get_list(range[0].value)
        check_entries_against_list(range, user_list)
    elif process_number == 9:
        limit = get_limit(range[0].value)
        check_character_limit(range, limit)
    elif process_number == 10:
        data_type = get_data_type(range[0].value)
        check_data_type(range, data_type)


def get_list(header):
    while (True):  # Loop until you get a valid text file
        print('-' * 40)
        list_file_path = input("Type path of the text file containing the list you would like " + header
                               + " checked against: ")
        list_file_path = os.path.abspath(list_file_path)
        if os.path.exists(list_file_path):
            if list_file_path[-4:] == ".txt":
                try:
                    user_list_file = open(list_file_path)
                    user_list = user_list_file.read().splitlines()
                    user_list_file.close()
                    break
                except Exception:
                    print("ERROR: Unable to open file.")
                    print()
            else:
                print("ERROR: Must be a .txt file.")
                print()
        else:
            print("ERROR: Invalid file path.")
            print()
    return user_list


def get_limit(header):
    while (True):  # Loop until you get a valid text file
        print('-' * 40)
        limit = input("Enter the character limit you would like used for " + header + ": ")
        try:
            limit = int(limit)
            break
        except Exception:
            print("ERROR: Must be a whole number.")
            print()
    return limit


def get_data_type(header):
    print('-' * 40)
    print("Select the data type option (0-" + str(len(DATA_TYPE_LIST) - 1) + ") you would like used for " + header)
    print('-' * 40)
    # Prints each item in list
    for item in DATA_TYPE_LIST:
        print("(" + str(DATA_TYPE_LIST.index(item)) + ") " + item)
    return get_user_selection(DATA_TYPE_LIST)


def clean_phone_number(range):
    # Setup regular expression
    phone_regex = re.compile(r'''(
        (?P<area_code>\d{3}|\((\s+)?\d{3}(\s+)?\)|\[(\s+)?\d{3}(\s+)?\])?       # Area code
        ((\s+)?(\s|-|\.)?(\s+)?)?                                               # Separator
        (?P<three_digits>\d{3})                                                 # First 3 digits
        ((\s+)?(\s|-|\.)?(\s+)?)?                                               # Separator
        (?P<four_digits>\d{4})                                                  # Last 4 digits
        (\s*(ext|x|ext.)\s*(?P<ext>\d{2,5}))?                                   # Extension
        )''', re.VERBOSE)

    strip_none_digits = re.compile(r'(\d+)')

    # Clean column
    for cell in range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Check against regex
        if phone_regex.search(str(cell.value)):
            match = phone_regex.search(str(cell.value))
            phone_number = ""
            if match.group('area_code'):
                area_code = strip_none_digits.search(match.group('area_code'))
                phone_number += "(" + area_code.group(0) + ") " + match.group('three_digits') + "-" \
                                + match.group('four_digits')
            else:
                phone_number += match.group('three_digits') + "-" + match.group('four_digits')
            if match.group('ext'):
                phone_number += 'x' + match.group('ext')
            cell.value = phone_number
        else:
            cell.value = ""


def clean_email_address(range):
    # Setup regular expression
    # Based email rules of information on this site:
    # https://help.returnpath.com/hc/en-us/articles/220560587-What-are-the-rules-for-email-address-syntax-
    email_regex = re.compile(r'''(
        ([a-zA-Z0-9](([a-zA-Z0-9!#$%&'*+/=?^_`{|.-]){,62}[a-zA-Z0-9])?)     # Recipient name
        (@)                                                                 # @ symbol
        ([a-zA-Z0-9](([a-zA-Z0-9.-]){,251}[a-zA-Z0-9])?)                    # Domain name
        (\.)                                                                # . symbol
        (com|org|net|edu|co)                                                # Top-level domain
        )''', re.VERBOSE)

    # Clean column
    for cell in range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Check against regex
        if email_regex.search(str(cell.value)):
            match = email_regex.search(str(cell.value))
            cell.value = match.group(1)
        else:
            cell.value = ""


def clean_states(range):
    # Setup regular expression
    remove_special_characters = re.compile(r'''[,.:;=?!"*%<>\-_(){}[\]\\]''')

    # Clean column
    for cell in range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Look for cell.value in states_lookup dictionary
        state = ((remove_special_characters.sub("", str(cell.value))).upper()).strip()
        if state in custom_dictionaries.states_lookup.keys():
            cell.value = custom_dictionaries.states_lookup[state]
        else:
            cell.value = ""


def clean_zip_codes(range):
    zip_regex = re.compile(r'''(
            (?P<five_digits>(\d)?(\d{4}))       # 5 digits
            ((\s+)?(-)(\s+)?)?                  # -
            (?P<four_digits>\d{4})?             # 4 digits
            )''', re.VERBOSE)

    # Clean column
    for cell in range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Check against regex
        if zip_regex.search(str(cell.value)):
            match = zip_regex.search(str(cell.value))
            # Pad five digits if needed
            number_of_zeros = 5 - len(str(match.group('five_digits')))
            zip_code = ("0" * number_of_zeros) + match.group('five_digits')
            if match.group('four_digits'):
                zip_code += "-" + match.group('four_digits')
            cell.value = zip_code
        else:
            cell.value = ""


def clean_dates(range):
    yyyy_mm_dd = re.compile(r'''(
                (?P<year>\d{4})                         # Year
                (-|/|\s)                                # Separator (- or / or ' ')
                (?P<month>(1[0-2])|[0]?[1-9])           # Month
                (-|/|\s)                                # Separator (- or / or ' ')
                (?P<day>(3[0-1])|[0]?[1-9]|[1-2][0-9])  # Day
                )''', re.VERBOSE)

    mm_dd_yyyy = re.compile(r'''(
                (?P<month>([0]?[1-9])|([1][0-2]))               # Month
                (-|/|\s)                                        # Separator (- or / or ' ')
                (?P<day>([0]?[1-9])|([1-2][0-9])|([3][0-1]))    # Day
                (-|/|\s)                                        # Separator (- or / or ' ')
                (?P<year>(\d{2})?(\d{2}))                       # Year
                )''', re.VERBOSE)

    month_word = re.compile(r'''(
                (?P<month>JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|JANUARY|FEBRUARY|MARCH|APRIL|JUNE|JULY|AUGUST
                          |SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)                 # Month
                (\s|\.)+                                                        # Separator
                (?P<day>([0]?[1-9])|([1-2][0-9])|([3][0-1]))                    # Day
                (\s|\.|,|')+                                                    # Separator
                (?P<year>(\d{2})?(\d{2}))                                       # Year
                )''', re.VERBOSE)

    # Clean column
    for cell in range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Check against regexes
        if yyyy_mm_dd.search(str(cell.value)):
            match = yyyy_mm_dd.search(str(cell.value))
            day = "0" * (2 - len(match.group('day'))) + match.group('day')
            month = "0" * (2 - len(match.group('month'))) + match.group('month')
            cell.value = match.group('year') + "-" + month + "-" + day

        elif mm_dd_yyyy.search(str(cell.value)):
            match = mm_dd_yyyy.search(str(cell.value))
            day = "0" * (2 - len(match.group('day'))) + match.group('day')
            month = "0" * (2 - len(match.group('month'))) + match.group('month')
            year = ""
            if int(match.group('year')) < 100:
                if int("20" + str(match.group('year'))) < (datetime.now()).year:
                    year += "20" + str(match.group('year'))
                else:
                    year += "19" + str(match.group('year'))
            else:
                year += str(match.group('year'))
            cell.value = year + "-" + month + "-" + day

        elif month_word.search(str(cell.value).upper()):
            match = month_word.search(str(cell.value).upper())
            day = "0" * (2 - len(match.group('day'))) + match.group('day')
            year = ""
            if int(match.group('year')) < 100:
                if int("20" + str(match.group('year'))) <= (datetime.now()).year:
                    year += "20" + str(match.group('year'))
                else:
                    year += "19" + str(match.group('year'))
            else:
                year += str(match.group('year'))
            month = custom_dictionaries.month_lookup[(str(match.group('month')).upper())[:3]]
            cell.value = year + "-" + month + "-" + day

        else:
            cell.value = ""


def clean_web_addresses(range):
    # Setup regular expression
    # Reference: https://www.regextester.com/93652
    web_address_regex = re.compile(r'''(
                        (http:\/\/www\.|https:\/\/www\.|http:\/\/|https:\/\/)?
                        [a-z0-9]+
                        ([\-\.]{1}[a-z0-9]+)*
                        \.
                        [a-z]{2,5}
                        (:[0-9]{1,5})?
                        (\/.*)?
                        )''', re.VERBOSE)

    # Clean column
    for cell in range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Check against regex
        if web_address_regex.search(str(cell.value)):
            match = web_address_regex.search(str(cell.value))
            cell.value = match.group(1)
        else:
            cell.value = ""


def clean_social_media(range):
    pass


def get_unique_entries(range):
    pass


def check_entries_against_list(range, l):
    # Convert list to uppercase so the check is case-insensitive
    for item in l:
        item = item.upper()

    # Clean column
    for cell in range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Look for cell.value in user_list, if not there, remove the entry in the Excel File
        if (str(cell.value)).upper() not in l:
            cell.value = ""


def check_character_limit(range, limit):
    # Clean column
    for cell in range:
        # Skip Header Row
        if cell.row == 1:
            continue
        # Check if cell.value is within character limit, if not, truncate
        if len(str(cell.value)) > limit:
            cell.value = (str(cell.value))[:limit]


def check_data_type(range, t):
    pass


init()
