'''
menus.py (Excel Command Line Tool)
Sydney Fowler and Matthew Hileman
15 December 2019
Description: Menus object, used to create menus throughout program.
'''

# ================ REFERENCES ================
# IMPORTLIB (needed import)

# ================ IMPORTS ================
# System
import os

# Custom
import menu_info

# Exterior
import importlib

# ================ NAVIGATIONAL CONSTANTS ================
# List of all menus
TOP_MENU_LIST = ["Analysis", "Cleanup", "Compress", "Duplicate Removal",
                 "Email", "Export", "Import", "Quit"]
ANALYSIS_MENU_LIST = ["Analyze new file", "Back to Main Menu", "Quit"]
CLEANUP_MENU_LIST = ["Clean new file", "Back to Main Menu", "Quit"]
COMPRESS_MENU_LIST = ["Compress new file", "Back to Main Menu", "Quit"]
DUPLICATE_MENU_LIST = ["Remove duplicate from new file", "Back to Main Menu", "Quit"]
EMAIL_MENU_LIST = ["Send an excel file via email", "Back to Main Menu", "Quit"]
EXPORT_MENU_LIST = ["Export new file", "Back to Main Menu", "Quit"]
IMPORT_MENU_LIST = ["Import new file", "Back to Main Menu", "Quit"]
QUIT_LIST = ["Yes, quit the program", "No, return to main menu"]


# Routes that each list item will flow to (file name)
TOP_MENU_ROUTE = ["analysis", "cleanup", "compress",
                  "duplicate_removal", "share",
                  "file_out", "file_in", "quit"]
ANALYSIS_MENU_ROUTE = ["analysis", "main", "quit"]
CLEANUP_MENU_ROUTE = ["cleanup", "main", "quit"]
COMPRESS_MENU_ROUTE = ["compress", "main", "quit"]
DUPLICATE_MENU_ROUTE = ["duplicate_removal", "main", "quit"]
EMAIL_MENU_ROUTE = ["share", "main", "quit"]
EXPORT_MENU_ROUTE = ["file_out", "main", "quit"]
IMPORT_MENU_ROUTE = ["file_in", "main", "quit"]
QUIT_ROUTE = ["quit", "main"]

# ================ MENU OBJECT ================
class Menu:

    # Initialization
    def __init__(self, name, list, route):
        self.min_option_value = 0
        self.max_option_value = len(list) - 1
        self.name = name
        self.list = list
        self.route = route
        self.selection_value = None
        self.selection_name = None

    # Prints menu options for user
    def print_menu_message(self):

        # Display print message from the import menu_info
        getattr(menu_info, "print_" + self.name + "_message")()

    # Gives menu for shiting menus
    def display_shift_menu(self):

        # Prints menu list
        self.print_menu_list()
        # Gets selection from user, stores if valid
        self.get_user_selection()
        # Routes to new file
        self.route_menu()

    def print_menu_list(self):
        # Tells user to select option
        print("Select an option (" + str(self.min_option_value) + "-" + str(self.max_option_value) + ")")
        print('-' * 40)
        # Prints each item in list
        for item in self.list:
            print("(" + str(self.list.index(item)) + ") " + str(item))

    # Gets input for menu selection and validates
    def get_user_selection(self):

        # Error checking loop - input is an integer and is a valid menu item
        while True:
            print('-' * 40)
            print("Choice: ", end = '')

            # Initialize choice
            user_choice = input()

            # Error handling: makes sure input is integer, stores interger
            try:
                user_choice = int(user_choice)
            # If input is not an integer, display error, has user try again.
            except ValueError:
                # Error Message
                print()
                print(str(user_choice) + " is not a valid input (NOT AN INT)")
                print("Choose a numeric value from the options above between (" +
                str(self.min_option_value) + "-" + str(self.max_option_value) + ").")
                continue

            # Error handling: makes sure the user's choice is a valid menu option
            if ( (user_choice < self.min_option_value)
            or (user_choice > self.max_option_value) ):
                # Error Message
                print()
                print(str(user_choice) + " is not a valid input.")
                print("Choose a numeric value from the options above between (" +
                str(self.min_option_value) + "-" + str(self.max_option_value) + ").")
                continue

            # If input is valid, return the input value, break from error loop
            else:
                self.selection_value = user_choice
                self.selection_name = self.route[user_choice]
                break

    # Directs user to respected file
    def route_menu(self):

        # Clears the screen (checks os and uses command for that system)
        os.system('cls' if os.name == 'nt' else 'clear')

        # Get import from predifined routes list
        imp = importlib.import_module(self.route[self.selection_value])

        # Decides to init or menu display
        target = None
        if self.name == self.selection_name:
            target = "init"
        else:
            target = "menu_header"

        # Go to routed import main
        getattr(imp, target)()


class Function_Menu(Menu):

    # Initialization
    def __init__(self, name, list, route):
        super().__init__(name, list, route)

    # Gives menu for shiting menus
    def display_shift_menu(self, *args):

        # Prints menu list
        self.print_menu_list()
        # Gets selection from user, stores if valid
        self.get_user_selection()
        # Routes to new file
        self.route_menu(args)

    # Prints menu options for user
    def print_menu_message(self):
        pass

    # Directs user to respected file
    def route_menu(self, args):

        # Clears the screen (checks os and uses command for that system)
        os.system('cls' if os.name == 'nt' else 'clear')

        imp = importlib.import_module(self.name)

        # Target is the selection user chose
        target = self.selection_name

        # Go to routed function
        getattr(imp, target)(args)


class Value_Menu(Menu):

    # Initialization
    def __init__(self, name, list, route):
        super().__init__(name, list, route)

    # Prints menu options for user
    def print_menu_message(self):
        pass

    # Gives menu for shiting menus
    def display_shift_menu(self):

        # Prints menu list
        self.print_menu_list()
        # Gets selection from user, stores if valid
        self.get_user_selection()
        # Routes selection value
        return self.route_menu()

    # Directs user to respected file
    def route_menu(self):

        # Clears the screen (checks os and uses command for that system)
        os.system('cls' if os.name == 'nt' else 'clear')

        # Target is the selection user chose
        value = self.selection_name

        # Return value
        return value
