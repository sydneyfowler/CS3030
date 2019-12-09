'''
compress.py (Excel Command Line Tool)
Sydney Fowler and Matthew Hileman
15 December 2019
Description: Allows the user to compress a file or folder.
'''

# ================ REFERENCES ================
# (none)

# ================ IMPORTS ================
# System
import zipfile
import os

# Custom
import menus
from excel_funcs import get_directory

# ================== SETUP ===================
def menu_header():
    # Print Main Compress Menu
    cleanup_main_menu = menus.Menu("compress", menus.COMPRESS_MENU_LIST, menus.COMPRESS_MENU_ROUTE)
    cleanup_main_menu.print_menu_message()
    cleanup_main_menu.display_shift_menu()

def init():

    # Get path and determine zipfile name
    path = os.path.abspath(get_directory([], "Type path of the file/folder you would like to zip: "))

    zipfile_name = path
    if zipfile_name.find(os.sep) != -1:
        zipfile_name = zipfile_name[(zipfile_name.rfind(os.sep)+1):]
    if zipfile_name.find('.') != -1:
        zipfile_name = zipfile_name[:zipfile_name.find('.')]
    zipped_file = zipfile.ZipFile(zipfile_name + ".zip", 'w')

    # If path is a directory, change directories to path and write all files in that directory to the zipfile
    if os.path.isdir(path):
        current_dir = os.curdir
        os.chdir(path)
        file_paths = os.listdir('.')
        for file in file_paths:
            zipped_file.write(file)

    # If path is a file, change directories to housing directory and write single file to the zipfile
    if os.path.isfile(path):
        current_dir = os.curdir
        os.chdir(os.path.dirname(path))
        zipped_file.write(os.path.basename(path))

    # Close the zipfile and move back to previous directory
    zipped_file.close()
    os.chdir(current_dir)

    print("Done! New file saved from " + path)
    input("Press enter to continue...")

    # Clears the screen (checks os and uses command for that system)
    os.system('cls' if os.name == 'nt' else 'clear')

    # Loop back to compress menu
    menu_header()

# For test purposes, will execute header if being run as main
if __name__ == '__main__':
    menu_header()
