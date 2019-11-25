"""
(Application Name)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
(Program Description)
"""
import os
import importlib
import menu_info
import file_in

'''-- CONSTANTS --'''
# List of all menus
TOP_MENU_LIST = ["Analysis", "Cleanup", "Compress", "Duplicate Removal",
                    "Email", "Export", "Import"]
ANALYSIS_MENU_LIST = []
CLEANUP_MENU_LIST = []
COMPRESS_MENU_LIST = []
DUPLICATE_MENU_LIST = []
EMIAL_MENU_LIST = []
EXPORT_MENU_LIST = []
IMPORT_MENU_LIST = [".scv", ".txt", ".pdf", ".py", ".rtf"]


TOP_MENU_ROUTE = ["anaysis.AnalysisMain", "cleanup.CleanupMain", "compress.CompressMain",
                    "duplicate_removal.DuplicateMain", "share.ShareMain",
                    "file_out.ExportMain", "file_in"]
ANALYSIS_MENU_ROUTE = []
CLEANUP_MENU_ROUTE = []
COMPRESS_MENU_ROUTE = []
DUPLICATE_MENU_ROUTE = []
EMIAL_MENU_ROUTE = []
EXPORT_MENU_ROUTE = []
IMPORT_MENU_ROUTE = ["todo", "todo", "todo", "todo", "todo"]

'''-- MENU CLASS --'''
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
    def PrintMenuMessage(self):

        # Display print message from the import menu_info
        getattr(menu_info, "Print" + self.name + "Message")()

    # Gives menu for shiting menus
    def DisplayShiftMenu(self):

        # Prints menu list
        self.PrintMenuList()
        # Gets selection from user, stores if valid
        self.GetUserSelection()
        # Routes to new file
        self.RouteMenu()

    def PrintMenuList(self):
        # Tells user to select option
        print("Select an option (" + str(self.min_option_value) + "-" + str(self.max_option_value) + ")")
        print('-' * 40)
        # Prints each item in list
        for item in self.list:
            print("(" + str(self.list.index(item)) + ") " + item)


    # Gets input for menu selection and validates
    def GetUserSelection(self):

        # Error checking loop - input is an integer and is a valid menu item
        while (True):
            print('-' * 40)
            print("Choice: ", end = '')

            # Initilize choice
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
                self.selection_name = self.list[user_choice]
                break

    # Directs user to respected file
    def RouteMenu(self):

        # Clears the screen (checks os and uses command for that system)
        os.system('cls' if os.name == 'nt' else 'clear')

        # Get import from predifined routes list
        imp = importlib.import_module(self.route[self.selection_value])

        # Go to routed import main
        getattr(imp, self.selection_name + "Main")()
