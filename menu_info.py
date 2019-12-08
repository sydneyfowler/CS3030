"""
(Application Name)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
(Program Description)
"""
import main
import menus

# ================ MENU MESSAGES ================
# Prints welcome message, purpose of program
def print_top_message():

    # Header/titles
    print()
    print(" Welcome to [Program Name]! ".center(70, "="))
    print('-' * 70)

    # Purpose statement, requirements
    print("The purpose of this utility program is to provide an easy method to")
    print("  work with excel files using python.")
    print("  IMPORTANT: This utility requires: os, openpyxl, pprint... (EDIT THIS)")
    print('-' * 70)

    # Insturctions
    print("First, choose an option below that you want to perform.")
    print("  You can choose an option for a detailed description of each.")
    print("  The menus will guide you through the desired process you want to")
    print("  perform.")
    print('-' * 70)
    print()

# Prints Import Information
def print_file_in_message():

    # Header
    print()
    print(" IMPORT TOOL ".center(70, "="))
    print('-' * 70)

    # Description
    print("The import tool will allow you to select a file and convert it to")
    print("  an excel file.")
    print('-' * 70)

    # Instructions
    print("  1st: Give the file to load (absolute or relative path).")
    print("    File types accepted: .csv")
    print()
    print("  2nd: Choose what to load from that file (a column, row, or entire file)")
    print()
    print("  3rd: Choose a file for the imported data")
    print("    (append, overwrite, or create new)")
    print('-' * 70)
    print()

# Prints Import Information
def print_file_out_message():

    # Header
    print()
    print(" OUTPUT TOOL ".center(70, "="))
    print('-' * 70)

    # Description
    print("The output tool will allow you to select an excel file and convert it")
    print("  to an csv.")
    print('-' * 70)

    # Instructions
    print("  1st: Give the file to export (absolute or relative path).")
    print("    File types accepted: .xlsx")
    print()
    print("  2rd: Choose a name for the exported data file")
    print("    (append, overwrite, or create new)")
    print('-' * 70)
    print()

# Prints Import Information
def print_share_message():

    # Header
    print()
    print(" SHARE (EMAIL) TOOL ".center(70, "="))
    print('-' * 70)

    # Description
    print("The share tool takes an excel file and sends it to a given set of")
    print("  emails contained within a .txt document.")
    print()
    print("  You will need: an excel file, a .txt containing emails, and an email ")
    print("  account from which to send the excel file.")
    print('-' * 70)

    # Instructions
    print("  1st: Give the file you wish to send (absolute or relative path).")
    print("    File types accepted: .xlsx")
    print()
    print("  2nd: Give the file with the emails you wish to send to.")
    print("    File types accepted: .txt")
    print()
    print("  3rd: Give your email credentials for the email account you wish")
    print("    to send from.")
    print('-' * 70)
    print()

# Prints Quit Message
def print_quit_message():

    # Header
    print()
    print(" CLEANUP FILE TOOL ".center(70, "="))
    print('-' * 70)

    # Description
    print("The cleanup tool takes an excel file and cleans it based on a list of")
    print("  options.")
    print()
    print("  You will need: an excel file and path to that excel file.")
    print('-' * 70)

    # Instructions
    print("  1st: Give the file you wish to cleanup (absolute or relative path).")
    print("    File types accepted: .xlsx")
    print()
    print("  2nd: Choose a column or row to cleanup.")
    print()
    print("  3rd: Choose an option to clean that selected row.")
    print('-' * 70)
    print()

# Prints Dluplicate Removal Message
def print_duplicate_removal_message():

    # Header
    print()
    print(" DUPLICATE REMOVAL TOOL ".center(70, "="))
    print('-' * 70)

    # Description
    print("The duplicate removal tool takes an excel file and removes extra rows/columns, ")
    print("  or duplicate items in rows/columns.")
    print()
    print("  You will need: an excel file and path to that excel file.")
    print('-' * 70)

    # Instructions
    print("  1st: Give the file you wish to remove duplicates (absolute or relative path).")
    print("    File types accepted: .xlsx")
    print()
    print("  2nd: Choose to cleanup rows, columns, a column, or row to remove duplicates.")
    print()
    print("  3rd: Output file will automatically save.")
    print('-' * 70)
    print()

# Prints Quit Message
def print_quit_message():

    # Header
    print()
    print(" QUIT MODULE ".center(70, "="))
    print('-' * 70)

    # Description
    print("Are you sure you want to terminate this tool?")
    print('-' * 70)
    print()
