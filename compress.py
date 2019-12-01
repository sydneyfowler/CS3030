'''
Final Project
Sydney Fowler and Matt Hileman
15-12-2019
Description: Allows the user to compress a file or folder.
'''

import zipfile
import os


def init():
    # Get path and determine zipfile name
    path = os.path.abspath(get_path())
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


def get_path():
    while (True):  # Loop until you get a valid Excel file
        path = input("Type path of the file/folder you would like to zip: ")
        if os.path.exists(path):
            break
        else:
            print("ERROR: Invalid path.")

    return path
