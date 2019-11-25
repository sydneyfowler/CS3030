"""
(Application Name)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
(Program Description)
"""
# Imports
import menus
import menu_info

if __name__ == '__main__':
    current_menu = menus.Menu("Top", menus.TOP_MENU_LIST, menus.TOP_MENU_ROUTE)
    current_menu.PrintMenuMessage()
    current_menu.DisplayShiftMenu()
