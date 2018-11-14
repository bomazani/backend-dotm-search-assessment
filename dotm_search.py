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
# import glob


# def find_dotm(path_to_search, text_to_search):
def find_dotm(path_to_search, text_to_find):
    """ iterate through dotm files, looking for text """
    # dotm_files = glob.iglob(folder + '/*.dotm')
    # strategy

    # create counters
    file_count = 0
    search_count = 0
    dotm_count = 0
    xml_count = 0
    match_count = 0
    found_count = 0
    open_count = 0
    char_count = 0
    # get a list of files to iterate over. maybe use os.listdir?
    dirs = os.listdir(path_to_search)
    # be sure to exclude any files not ending with ".dotm"
    for file in dirs:
        file_count += 1
        get_ext = os.path.splitext(file)
        if get_ext[1] == '.dotm':
            # print(path_to_search)
            dotm_count += 1
            fill_path = os.path.join(path_to_search, file)
            # print(fill_path)
            # open zip files like this:
            # with zipfile.ZipFile(fill_path) as z:
            with zipfile.ZipFile(fill_path, "r") as z:
                # print('***************')
                # Get z's table of contents:
                names = z.namelist()
                found_files = list(z.namelist())
                for file in found_files:
                    found_count += 1
                    # print('found file = ' + str(file))
                xml_files = list(
                    filter(lambda x: x.endswith('.xml'), z.namelist()))
                # print(names)
                # print('---------------')
                # print(xml_files)
                for file in xml_files:
                    xml_count += 1
                    found_match = False
                    with z.open(file, "r") as xml:
                        content = xml.read()

                        open_count += 1
                        for index, search_item in enumerate(content):
                            seperator = ''
                            if search_item == '$':
                                found_match = True
                                # print('Search_Item = ' + search_item)
                                # print('index = ' + str(index))
                                found_item = content[index-40:index+40]
                                found_string = ''
                                for i in found_item:
                                    found_string += i
                                print('Match found in file ' + str(fill_path))
                                print('   ...' + found_string + '...')
                        if found_match:
                            match_count += 1
                            print(' ')
                


    # names = z.namelist()

    # using z again, open the 'word/document.xml' file
    # read the z file line by line, and search for text_to_search
    # if you find a line containing the desired text, print it to console
    # update your match_count and search_count appropriately
    # print('file_count = ' + str(file_count))
    # print('search_count = ' + str(search_count))
    # print('found_count = ' + str(found_count))
    print('Total dotm files searched: ' + str(dotm_count))
    print('Total dotm files matched: ' + str(match_count))

    # print('word_doc_count = ' + str(word_doc_count))
    # print('xml_count = ' + str(xml_count))
    # print('open_count = ' + str(open_count))


def create_parser():
    # create my own instance of a parcer
    parser = argparse.ArgumentParser()
    # add my own arguments that I expect on the cmd line
    parser.add_argument('text', help='text to search for dotm files')
    parser.add_argument(
        '-d', '--dir', help='directory of dotm files to search', default='.')
    return parser


def main():
    parser = create_parser()
    # running the parser produces a 'namespace' of items.
    # namespace is really a dict.
    my_args = parser.parse_args()
    if not my_args:
        parser.print_usage()
        sys.exit(1)

    text_to_find = my_args.text
    path_to_search = my_args.dir
    find_dotm(path_to_search, text_to_find)

    # print(my_args.dir)
    # print(text_to_find)
    # print(my_args.text)
    # print(path_to_search)


if __name__ == '__main__':
    main()
