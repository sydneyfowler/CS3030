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

CLEANUP_OPTIONS_LIST = ["Cleanup Phone Numbers", "Cleanup Email Addresses", "Cleanup States", "Cleanup Zip Codes",
                        "Cleanup Dates", "Cleanup Web Address", "Cleanup Social Media",
                        "Produce List of Unique Entries", "Check Entries Against List", "Truncate to Character Limit",
                        "Check Data Type", "Finish Sheet"]


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
            user_selection = get_user_selection()
            if user_selection == (len(CLEANUP_OPTIONS_LIST) - 1): # Check if user wants to break out of sheet
                break
            else:
                sheet_header_lookup[sheet_name][header] = user_selection

    # Process data
    for sheet_name in sheets:
        sheet = wb.get_sheet_by_name(sheet_name)
        for col in range(1, sheet.max_column):
            process_number = sheet_header_lookup[sheet_name][sheet.cell(row=1, column=col).value]
            if process_number is not None:
                process_column(sheet[get_column_letter(col)], int(process_number))

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


def get_user_selection():
    # Error checking loop - input is an integer and is a valid menu item
    while (True):
        print('-' * 40)
        print("Choice: ", end='')

        # Initilize choice
        user_choice = input()

        # Error handling: makes sure input is integer, stores interger
        try:
            user_choice = int(user_choice)
        # If input is not an integer, display error, has user try again.
        except ValueError:
            # Error Message
            print()
            print(str(user_choice) + " is not a valid input (NOT AN INT)")
            print("Choose a numeric value from the options above between (0-" + str(len(CLEANUP_OPTIONS_LIST) - 1)
                  + ").")
            continue

        # Error handling: makes sure the user's choice is a valid menu option
        if (user_choice < 0) or (user_choice >= len(CLEANUP_OPTIONS_LIST)):
            # Error Message
            print()
            print(str(user_choice) + " is not a valid input.")
            print("Choose a numeric value from the options above between (0-" + str(len(CLEANUP_OPTIONS_LIST) - 1)
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
        check_entries_against_list(range)
    elif process_number == 9:
        check_character_limit(range, 10)
    elif process_number == 10:
        check_data_type(range, "int")


def clean_phone_number(range):
    # Setup regular expression
    phone_regex = re.compile(r'''(
        (\d{3}|\(\d{3}\))?                  # Area code
        (\s|-|\.)?                          # Separator
        \d{3}                               # First 3 digits
        (\s|-|\.)                           # Separator
        \d{4}                               # Last 4 digits
        (\s*(ext|x|ext.)\s*\d{2,5})?        # Extension
        )''', re.VERBOSE)


def clean_email_address(range):
    # Setup regular expression
    # Based email rules of information on this site:
    # https://help.returnpath.com/hc/en-us/articles/220560587-What-are-the-rules-for-email-address-syntax-
    email_regex = re.compile(r'''(
        ([a-zA-Z0-9](([a-zA-Z0-9!#$%&'*+/=?^_`{|.-]){,62}[a-zA-Z0-9])?)     # Recipient name
        (@)                                                                 # @ symbol
        ([a-zA-Z0-9](([a-zA-Z0-9.-]){,251}[a-zA-Z0-9])?)                    # Domain name
        (\.)                                                                # . symbol
        (com|org|net)                                                       # Top-level domain
        )''', re.VERBOSE)


def clean_states(range):
    pass


def clean_zip_codes(range):
    zip_regex = re.compile(r'''(
            (\d{5})                             # 5 digits
            (-.)?                               # -
            (\d{4})?                            # 4 digits
            )''', re.VERBOSE)


def clean_dates(range):
    yyyy_mm_dd = re.compile(r'''(
                (\d{4})                         # Year
                (-|/)                           # Separator (- or /)
                ((1[0-2])|0[1-9])               # Month
                (-|/)                           # Separator (- or /)
                ((3[0-1])|0[1-9]|[1-2][0-9])    # Day
                )''', re.VERBOSE)

    mm_dd_yyyy = re.compile(r'''(
    (([0][1-9])|([1][0-2]))                 # Month
    (-|/)                                   # Separator (- or /)
    (([0][1-9])|([1-2][0-9])|([3][0-1]))    # Day
    (-|/)                                   # Separator (- or /)
    ((\d{2})?(\d{2}))                       # Year
    )''', re.VERBOSE)


def clean_web_addresses(range):
    pass


def clean_social_media(range):
    pass


def get_unique_entries(range):
    pass


def check_entries_against_list(range):
    pass


def check_character_limit(range, limit):
    pass


def check_data_type(range, t):
    pass


init()