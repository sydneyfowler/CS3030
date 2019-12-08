"""
Output (file_out)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
Description: Converts a .xlsx to .csv
"""

import pandas
from excel_funcs import get_directory
import menus


# ================ REFERENCES ================
# PANDAS - (NEED XLRD installed)

# ================== SETUP ===================

def menu_header():

    # Print Import Message Above
    output_menu = menus.Menu("file_out", menus.EXPORT_MENU_LIST, menus.EXPORT_MENU_ROUTE)
    output_menu.PrintMenuMessage()
    output_menu.DisplayShiftMenu()

def init():

    # Get input for excel file
    wb_path = get_directory([".xlsx"], "Type path of your excel file (.xlsx): ")
    wb_xls = pandas.read_excel(wb_path, 'Sheet1', index_col = None)

    # Get input where to save export file (.csv file)
    export_path = os.path.dirname(os.path.abspath(wb_path))
    export_path += "/" + input("Input new file's name (saves to same directory): ") + ".csv"
    wb_xls.to_csv(export_path, encoding = 'utf-8', index = False)

    # Success message
    print("Success!\n" + wb_path + " was saved to " + export_path)
    input("Press enter to continue...")

    # Display Menu Header again.
    menu_header()
