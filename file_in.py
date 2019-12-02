"""
(Application Name)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
(Program Description)
"""

import os
import sys
import re
import pandas

import menus


# ================ REFERENCES ================
# PANDAS - https://pandas.pydata.org/ (NEED XLRD)

# ================== SETUP ===================

def menu_header():

    # Print Import Message Above
    import_menu = menus.Menu("file_in", menus.IMPORT_MENU_LIST, menus.IMPORT_MENU_ROUTE)
    import_menu.PrintMenuMessage()
    import_menu.DisplayShiftMenu()

def init():

    # Get input for .csv file
    wb_path = get_directory([".csv"], "Type path of your .csv file: ")
    wb_csv = pandas.read_csv(wb_path)

    # Get input for where to save new excel file
    export_path = get_directory([], "Type path of your desired conversion (don't include file name): ")
    export_path += "/" + input("Name the new file (don't need to include file type): ") + ".xlsx"
    wb_csv.to_excel(export_path, index = False)

    # Sucess message
    print("Success!\n" + wb_path + " was saved to " + export_path)
    input("Press enter to continue...")

    # Display Menu Header again
    menu_header()

# Returns imported file
def get_directory(type_array, message):

    error_found = 0
    while (True):

        # Import path input
        file_path = input(message)

        # Check that file exists
        if os.path.exists(file_path):
            for type in type_array:
                # Checks for correct file type
                if ( (file_path[-len(type):] != type) ):
                    error_found = 1
                else:
                    error_found = 0

            if (error_found):
                print("ERROR: Invalid file TYPE. Must be " + str(type_array))
                print()
                continue
            else:
                print('-' * 40)
                print()
                break

        else:
            print("ERROR: Invalid file PATH.")
            print()
            continue

    return file_path
