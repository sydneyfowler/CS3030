"""
(Application Name)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
(Program Description)
"""
import main
import menus

'''-- WELCOME MESSAGE --'''
# Prints welcome message, purpose of program
def PrintTopMessage():

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
def PrintImportMessage():

    # Header
    print()
    print(" IMPORT MODULE ".center(70, "="))
    print('-' * 70)

    # Insutrctions
    print("The import tool will allow you to select a file and convert it to")
    print("  an excel file. You can choose to import from rows, columns, or")
    print("  enitre files from there. Fist, give the file to load (absolute or ")
    print("  relative path). File types accepted: .csv .txt .pdf .py .rtf")
    print('-' * 70)
    print()
