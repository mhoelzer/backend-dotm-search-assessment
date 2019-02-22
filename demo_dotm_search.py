#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "testdemo"

import os
import zipfile


def main(directory, text_to_search):
    file_list = os.listdir(directory)
    files_searched = 0
    files_matched = 0
    print("Searching directory {} for text '{}' ...".format(directory, text_to_search))
    for file in file_list:
        files_searched += 1
        full_path = os.path.join(directory, file)
        if not full_path.endswith(".dotm"):
            print("this isn't a dotm {}".format(full_path))
            continue
        if not zipfile.is_zipfile(full_path):
            print("this isn't a zip file {}".format(full_path))
            continue
        with zipfile.ZipFile(full_path, "r") as zipped:
            toc = zipped.namelist()
            if "word/document.xml" in toc:
                with zipped.open("word/document.xml", "r") as doc:
                    for line in doc:
                        i = line.find(text_to_search)
                        if i >= 0:
                            files_matched += 1
                            print("  ...{}...".format(line[i - 40:i + 40]))
    print("Files matched: {}".format(files_matched))
    print("Files searched: {}".format(files_searched))
                            

if __name__ == '__main__':
    main("dotm_files", "$")
