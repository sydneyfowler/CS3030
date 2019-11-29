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
    print(" IMPORT MODULE ".center(70, "="))
    print('-' * 70)

    # Description
    print("The import tool will allow you to select a file and convert it to")
    print("  an excel file. You can choose to import from rows, columns, or")
    print("  enitre files from there.")
    print('-' * 70)

    # Instructions
    print("  1st: Give the file to load (absolute or relative path).")
    print("    File types accepted: .csv .txt .pdf .py .rtf")
    print()
    print("  2nd: Choose what to load from that file (a column, row, or entire file)")
    print()
    print("  3rd: Choose a file for the imported data")
    print("    (append, overwrite, or create new)")
    print('-' * 70)
    print()

# Prints Import Information
def print_share_message():

    # Header
    print()
    print(" SHARE (EMAIL) ".center(70, "="))
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
