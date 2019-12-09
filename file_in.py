'''
file_in.py (Excel Command Line Tool)
Sydney Fowler and Matthew Hileman
15 December 2019
Description: Converts a .csv to .xlsx
'''

# ================ REFERENCES ================
# PANDAS - (NEED XLRD installed)

# ================ IMPORTS ================
# System
import os
import sys

# Custom
from excel_funcs import get_directory
import menus

# Exterior
import pandas

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
    export_path = os.path.dirname(os.path.abspath(wb_path))
    export_path += "/" + input("Input new file's name (saves to same directory): ") + ".xlsx"
    wb_csv.to_excel(export_path, index = False)

    # Sucess message
    print("Done! New file saved to " + export_path)
    input("Press enter to continue...")

    # Display Menu Header again
    menu_header()
