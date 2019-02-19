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

# open dotm fileand earch text ; $; dotmfiles; rerad binary


def main(directory, text_to_search):
    # os listdir to get outer then go in for .m
    # dotm is a zip
    # ptyhon zip file tutorrial
    # if $ in file...
    dir_path = os.listdir(directory)
    files_searched = 0
    files_matched = 0
    with zipfile.ZipFile(dir_path, "r") as zipped_file:
        if dir_path.endswith(".dotm"):
            with zipped_file.open("word/document.xml", "r") as opened_file:
                for file in dir_path:
                    print(file)
    print("Searching directory {} for text \"{}\" ...".format(
        dir_path, text_to_search))
    print("Files searched: {}".format(files_searched))
    print("Files matched: {}".format(files_matched))


def create_parser():
    parser = argparse.ArgumentParser(description="Process files to find dotm")
    parser.add_argument("--dir", help="directory for searching")
    parser.add_argument("text_to_search", help="text for searching")
    return parser


if __name__ == "__main__":
    parser = create_parser()
    args = parser.parse_args()
    main(args.directory, args.text_to_search)