#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the "word/document.xml" section of a MSWord .dotm file.
"""
__author__ = "mhoelzer"


import argparse
import os
import zipfile
import sys
# open dotm file and search text; $; dotmfiles; read binary


def main(directory, text_to_search):
    """
    searching and printing only the counts and matching text
    because we have to follow the rules like nerds
    """
    # os listdir to get outer then go in for .m
    # dotm is a zip
    # ptyhon zip file tutorrial
    # if $ in file...
    file_list = os.listdir(directory)  # all files in dir
    files_searched = 0
    files_matched = 0
    print("Searching directory {} for text \"{}\" ...".format(
        directory, text_to_search))
    for file in file_list:
        if file.endswith(".dotm"):
            files_searched += 1
            # join dir and file names
            full_path = os.path.join(directory, file)
            # is it a dotm?
            if not full_path.endswith(".dotm"):
                print("File is not .dotm: {}. Do you not understand the point of this code?".format(
                    full_path))
                continue
            # is it a zip?
            if not zipfile.is_zipfile(full_path):
                print("File is not zip: {}. Do you not understand the point of this code?".format(
                    full_path))
                continue
            with zipfile.ZipFile(full_path, "r") as zipped_file:
                table_of_contents = zipped_file.namelist()
                if "word/document.xml" in table_of_contents:
                    with zipped_file.open("word/document.xml", "r") as opened_file:
                        for line in opened_file:
                            # if text_to_search in line:
                            #     print("Match found in file {}".format(full_path))
                            index_num = line.find(text_to_search)
                            if index_num >= 0 and index_num in range(len(line)):
                                files_matched += 1
                                print("Match found in file {}".format(full_path))
                                # check for limits on  l and r
                                sliced_line = line[index_num -
                                                   40:index_num + 40]
                                print("   ...{}...".format(sliced_line))
    print("Total dotm files searched: {}".format(files_searched))
    print("Total dotm files matched: {}".format(files_matched))


# having this separate frorm the if name... makes things easier to test
def create_parser():
    parser = argparse.ArgumentParser(description="Process files to find dotm")
    # ^^^ gives us the parser obj
    parser.add_argument("--dir", help="directory for searching",
                        default=".")  # . means current dir
    parser.add_argument("text_to_search", help="text for searching")
    return parser


if __name__ == "__main__":
    parser = create_parser()
    # if not namespace:
    if len(sys.argv) <= 1:
        parser.print_usage()
        exit(1)
    namespace = parser.parse_args()
    main(namespace.dir, namespace.text_to_search)
    # ^^^ this is accessed by how we named/defined it in the .add_arg
    # main("--dir", "$")  # simple test
