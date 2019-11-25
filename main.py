"""
(Application Name)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
(Program Description)
"""
# Imports
import menus
import menu_info


def menu_header():
    main_menu = menus.Menu("top", menus.TOP_MENU_LIST, menus.TOP_MENU_ROUTE)
    main_menu.PrintMenuMessage()
    main_menu.DisplayShiftMenu()


# ================ TOP MENU ================
if __name__ == '__main__':
    menu_header()
