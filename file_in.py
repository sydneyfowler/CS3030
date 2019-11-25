"""
(Application Name)
Created by: Matthew Hileman & Sidney Fowler, 20 November 2019
(Program Description)
"""
import main
import menus
import menu_info

def ImportMain():

    # Print Import Message Above
    import_menu = menus.Menu("Import", menus.IMPORT_MENU_LIST, menus.IMPORT_MENU_ROUTE)
    import_menu.PrintMenuMessage()

    # PUT IMPORT FUNCTIONALITY HERE #

    #main.MainLoop(menus.Menu(menus.IMPORT_MENU_LIST, menus.IMPORT_MENU_ROUTE), 'import')
