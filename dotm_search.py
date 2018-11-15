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


def find_dotm(path_to_search, text_to_find):
    """ iterate through dotm files, looking for text """
    dotm_count = 0
    match_count = 0
    dirs = os.listdir(path_to_search)
    for file in dirs:
        get_ext = os.path.splitext(file)
        if get_ext[1] == '.dotm':
            dotm_count += 1
            file_path = os.path.join(path_to_search, file)
            with zipfile.ZipFile(file_path, "r") as unzipped:
                with unzipped.open('word/document.xml', "r") as xml:
                    content = xml.read()
                    match_count += search_for_string(content, file_path)            
    print('Total dotm files searched: ' + str(dotm_count))
    print('Total dotm files matched: ' + str(match_count))
    

def search_for_string(content, file_path):
    for index, search_item in enumerate(content):
        if search_item == '$':
            print('Match found in file ' + str(file_path))
            print('   ...' + content[index-40:index+40] + '...')
            return 1
    return 0

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
