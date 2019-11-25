"""
(Application Name)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
(Program Description)
"""
import main
import menus
import menu_info

def menu_header():

    # Print Import Message Above
    import_menu = menus.Menu("file_in", menus.IMPORT_MENU_LIST, menus.IMPORT_MENU_ROUTE)
    import_menu.PrintMenuMessage()
