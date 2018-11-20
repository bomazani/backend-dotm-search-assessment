#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.

author: bobh
"""

# imports go at the top of your file, after the module docstring above.
import sys
import os
import argparse
import zipfile
from itertools import count


def find_dotm(path_to_search, text_to_find):
    """ iterate through directory, unzip dotm files & search for word/document.xml files """
    print("Searching directory {} for text {} ...".format(
        path_to_search, text_to_find))
    dotm_count = 0
    match_count = 0
    total_count = 0
    dirs = os.listdir(path_to_search)
    for file in dirs:
        get_ext = os.path.splitext(file)
        if get_ext[1] == '.dotm':
            dotm_count += 1
            file_path = os.path.join(path_to_search, file)
            with zipfile.ZipFile(file_path, "r") as unzipped:
                with unzipped.open('word/document.xml', "r") as xml:
                    content = xml.read()
                    match_list = search_for_string(content, text_to_find)
                    if match_list:
                        match_count += 1
                        total_count += len(match_list)
                        print('Match found in file ' + str(file_path))
                        for m in match_list:
                            print('   ...' + content[m-40:m+40] + '...')
    print('Total dotm files searched: ' + str(dotm_count))
    print('Total dotm files matched: ' + str(match_count))
    print('Total dotm files : ' + str(total_count))


def search_for_string(content, search_text):
    results = []
    start_pos = 0
    while True:
        index = content.find(search_text, start_pos)
        if index < 0:
            return results
        results.append(index)
        start_pos = index + 1


def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('text', help='text to search for dotm files')
    parser.add_argument(
        '-d', '--dir', help='directory of dotm files to search', default='.')
    return parser


def main():
    parser = create_parser()
    my_args = parser.parse_args()
    if not my_args:
        parser.print_usage()
        sys.exit(1)

    text_to_find = my_args.text
    path_to_search = my_args.dir
    find_dotm(path_to_search, text_to_find)


if __name__ == '__main__':
    main()
