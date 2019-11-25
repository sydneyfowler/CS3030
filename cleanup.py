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


def init():
    wb_path = get_wb_path()
    wb = openpyxl.load_workbook(wb_path)
    sheets = wb.get_sheet_names()

    for sheet_name in sheets:
        pass

    # Save to a copy of the workbook
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


def clean_phone_number():
    # Setup regular expression
    phone_regex = re.compile(r'''(
        (\d{3}|\(\d{3}\))?                  # Area code
        (\s|-|\.)?                          # Separator
        \d{3}                               # First 3 digits
        (\s|-|\.)                           # Separator
        \d{4}                               # Last 4 digits
        (\s*(ext|x|ext.)\s*\d{2,5})?        # Extension
        )''', re.VERBOSE)


def clean_email_address():
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


def clean_states():
    pass


def clean_zip_codes():
    zip_regex = re.compile(r'''(
            (\d{5})                             # 5 digits
            (-.)?                               # -
            (\d{4})?                            # 4 digits
            )''', re.VERBOSE)


def clean_dates():
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


def clean_web_addresses():
    pass


def clean_social_media():
    pass


def get_unique_entries():
    pass


def check_entries_against_list():
    pass


def check_character_limit(limit):
    pass


def check_data_type(t):
    pass
