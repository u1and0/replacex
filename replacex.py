#!/usr/bin/env python3
"""Replace text and highlight it in docx files.
"""
import sys
import argparse
from itertools import chain
from docx import Document
# from docx.shared import Pt
from docx.shared import RGBColor

VERSION = 'v1.0.1'
CRED = '\033[91m'
CEND = '\033[0m'


def replace_text(paragraph, before, after):
    """paragraph内の文字列beforeをafterへ置換する"""
    replaced_text = paragraph.text.replace(before, after)
    if paragraph.text != replaced_text:
        paragraph.text = replaced_text
        # Modify docx sentence
        # paragraph.runs[0].font.size = Pt(10.5)
        paragraph.runs[0].font.color.rgb = RGBColor(235, 0, 0)
        yield paragraph.text


def main(old, new, *filenames, dryrun=False, verbose=False):
    """execute replace_text to multiple files"""
    for filename in filenames:
        document = Document(filename)
        if verbose or dryrun:
            print("==filename:", filename, "==")
        # Rewrite sentence
        sentence_paragraphs = (paragraph for paragraph in document.paragraphs)
        # Rewrite table
        table_paragraphs = (paragraph for table in document.tables
                            for row in table.rows for cell in row.cells
                            for paragraph in cell.paragraphs)
        # Concat iter
        paragraphs = chain(sentence_paragraphs, table_paragraphs)
        # Edit contents
        for paragraph in paragraphs:
            text_it = replace_text(paragraph, old, new)
            for text in text_it:
                # Print out result to stdout if verbose mode or dryrun mode
                if verbose or dryrun:
                    colored = text.replace(new, CRED + new + CEND)
                    print(colored)
        # Save Document unless dryrun mode
        if not dryrun:
            document.save(filename)


def parse():
    """arg parser"""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('old', type=str, help='old word')
    parser.add_argument('new', type=str, help='new word')
    parser.add_argument('files', type=str, nargs='*', help='docx file path')
    parser.add_argument(
        '-n',
        '--dryrun',
        help='DO NOT save docx file just print replacement result.',
        action='store_true',
        default=False,
    )
    parser.add_argument(
        '-v',
        '--verbose',
        help='print replacement result to stdout',
        action='store_true',
        default=False,
    )
    parser.add_argument('-V', '--version', action='store_true')
    return parser.parse_args()


if __name__ == '__main__':
    argv = parse()
    if argv.version:
        print('replacex:', VERSION)
        sys.exit(0)
    main(argv.old,
         argv.new,
         *argv.files,
         dryrun=argv.dryrun,
         verbose=argv.verbose)
