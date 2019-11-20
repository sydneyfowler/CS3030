"""
(Application Name)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
(Program Description)
"""
# Imports
import os
import Import

'''-- CONSTANTS --'''
# List of all options the user can perform at startup (main menu)
MAIN_OPTIONS_LIST = ["Exit", "Analysis", "Cleanup", "Compress",
                    "Duplicate Removal", "Email", "Export", "Import"]
# Min and Max menu inputs accepted
MIN_OPTION_VALUE = 0
MAX_OPTION_VALUE = len(MAIN_OPTIONS_LIST) - 1

'''-- FUNCTIONS --'''
# Prints welcome message, purpose of program
def PrintWelcomeMessage():

    # Header/titles
    print()
    print(" Welcome to [Program Name]! ".center(70, "="))
    print('-' * 70)

    # Purpose statement, requirements
    print("The purpose of this utility program is to provide an easy method to")
    print("work with excel files using in python (EDIT THIS)")
    print("IMPORTANT: This utility requires: os, openpyxl, pprint... (EDIT THIS)")
    print('-' * 70)

    # Purpose statement, requirements
    print("First, choose an option below that you want to perform.")
    print("You can choose an option for a detailed description of each.")
    print("The menus will guide you through the desired process you want to perform.")
    print('-' * 70)

    print()

# Prints options, based on a list
def PrintMenuList(menu_list):

    # Tells user to select option
    print("Select an option (" + str(MIN_OPTION_VALUE) + "-" + str(MAX_OPTION_VALUE) + ")")
    print('-' * 20)

    # Prints each item in list
    for item in menu_list:
        print("(" + str(menu_list.index(item)) + ") " + item)


# Gets input from user, validates it as acceptable menu option, returns input
def GetUserSelection():

    # Error checking loop - input is an integer and is a valid menu item
    while (True):
        print("\nChoice: ", end = '')

        # Error handling: makes sure input is integer, stores interger
        try:
            user_choice = int(input())

        # If input is not an integer, display error, has user try again.
        except ValueError:

            # Error Message
            print(str(user_choice) + " is not a valid input. Choose a numeric value between (" +
            str(MIN_OPTION_VALUE) + "-" + str(MAX_OPTION_VALUE) + "): ")
            continue

        # Error handlin: makes sure the user's choice is a valid menu option
        if ( (user_choice < MIN_OPTION_VALUE)
        or (user_choice > MAX_OPTION_VALUE) ):

            # Error Message
            print(str(user_choice) + " is not a valid input. Choose a numeric value between (" +
            str(MIN_OPTION_VALUE) + "-" + str(MAX_OPTION_VALUE) + "): ")
            continue

        # If input is valid, return the input value, break from error loop
        else:
            return user_choice
            break

'''--- MAIN ---'''
previous_menu = 0
current_menu = "Main Menu"

# Prints welcome
PrintWelcomeMessage()

# Prints menu list
PrintMenuList(MAIN_OPTIONS_LIST)

# Stores menu selection if valid
user_selection = MAIN_OPTIONS_LIST[GetUserSelection()]

# Clears the screen (checks os and uses command for that system)
os.system('cls' if os.name == 'nt' else 'clear')

# Depending on selection, runs function for that selection from other file
# IMPORTANT: ONLY WORKS WITH IMPORT RIGHT NOW, NEED TO STANDARDIZE TO WORK WITH ALL MENU OPTIONS.
# COMMENT OUT BELOW LINE TO TEST MAIN WITH ANY SELECTION, LEAST FUNCTION MAY NOT EXIST.
getattr(Import, user_selection + "Main")()

'''
Put other functions here
'''
