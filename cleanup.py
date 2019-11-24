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
    pass


def clean_email_address():
    pass


def clean_states():
    pass


def clean_zip_codes():
    pass


def clean_dates():
    pass


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
