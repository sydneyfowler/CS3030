"""
Quit Tool
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
Description: Quits the program after displaying confirmation.
"""
import menus
import sys

# ================== SETUP ===================
def menu_header():

    # Print Import Message Above
    import_menu = menus.Menu("quit", menus.QUIT_LIST, menus.QUIT_ROUTE)
    import_menu.PrintMenuMessage()
    import_menu.DisplayShiftMenu()

def init():
    sys.exit()
