'''
excel_funcs.py (Excel Command Line Tool)
Sydney Fowler and Matthew Hileman
15 December 2019
Description: Common excel related methods used throughout several tools.
            Includes: get_directory, save_file, get_sheet
'''

# ================ IMPORTS ================
# System
import os

# Custom
import menus

# ================ METHODS ================
# Get file function, used in various tools
def get_directory(type_array, message):

    print()
    print(" Input a file ".center(70, "="))
    print('-' * 70)

    # Initialize error flag
    error_found = 0
    while True:

        # Import path input
        file_path = input(message)

        # Check that file exists
        if os.path.exists(file_path):
            for type in type_array:

                # Checks for correct file type
                if file_path[-len(type):] != type:
                    error_found = 1
                else:
                    error_found = 0

            # If incorrect file type, print error, loop.
            if error_found:
                print("ERROR: Invalid file TYPE. Must be " + str(type_array))
                print()
                continue
            else:
                print('-' * 40)
                print()
                break

        # If incorrect path or file does not exist, print error, loop.
        else:
            print("ERROR: Invalid file PATH.")
            print()
            continue

    # Happy path, returns file and file path. Will return path if no file.
    return file_path


# Saves copy of a wb
def save_file(wb, wb_path, type):

    print()
    print(" Enter a name for new save file ".center(70, "="))
    print('-' * 70)

    # Removes original file name
    save_path = os.path.dirname(os.path.abspath(wb_path))

    # Gets name for new file
    save_path += "/" + input("Input new file's name (saves to same directory): ") + type

    # Save to a new copy of the workbook
    wb.save(save_path)
    print("Done! New file saved to " + save_path)
    input("Press enter to continue...")


# Has user select sheet to perform actions on
def get_sheet(wb):

    # Creates menu
    sheets = wb.sheetnames      # Edited depreciated function: "wb.get_sheet_names()"
    sheet_menu = menus.Value_Menu("duplicate_removal", sheets, sheets)
    print()
    print(" Choose a sheet to remove duplicates ".center(70, "="))
    print('-' * 70)

    # User selects sheet to use, returns that sheet
    sheet = wb.get_sheet_by_name(sheet_menu.display_shift_menu())
    return sheet
