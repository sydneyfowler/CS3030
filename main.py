'''
main.py (Excel Command Line Tool)
Sydney Fowler and Matthew Hileman
15 December 2019
Description: Used as a launcher. Starting point, top menu.
'''

# ================ IMPORTS ================
# Custom
import menus
import menu_info

# ================ TOP MENU ================
def menu_header():
    main_menu = menus.Menu("top", menus.TOP_MENU_LIST, menus.TOP_MENU_ROUTE)
    main_menu.print_menu_message()
    main_menu.display_shift_menu()


# If being run as main, display top menu
if __name__ == '__main__':
    menu_header()
