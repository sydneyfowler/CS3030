'''
quit.py (Excel Command Line Tool)
Sydney Fowler and Matthew Hileman
15 December 2019
Description: Quits the program after displaying confirmation.
'''
# ================ REFERENCES ================
# (none)

# ================ IMPORTS ================
# System
import sys

# Custom
import menus

# ================== SETUP ===================
def menu_header():

    # Print Import Message Above
    import_menu = menus.Menu("quit", menus.QUIT_LIST, menus.QUIT_ROUTE)
    import_menu.print_menu_message()
    import_menu.display_shift_menu()

def init():
    sys.exit()
