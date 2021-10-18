#!/usr/bin/env python3
"""Replace text and highlight it in docx files.

usage:
    $ python replacex.py OLDWORD NEWWORD [FILENAMES...]
"""
import sys
import argparse
from docx import Document
# from docx.shared import Pt
from docx.shared import RGBColor

VERSION = 'v0.2.0'
CRED = '\033[91m'
CEND = '\033[0m'


def replace_text(paragraph, before, after):
    """paragraph内の文字列beforeをafterへ置換する"""
    replaced_text = paragraph.text.replace(before, after)
    if paragraph.text != replaced_text:
        paragraph.text = replaced_text
        # Print out result
        colored = paragraph.text.replace(after, CRED + after + CEND)
        print(colored)
        # Modify docx sentence
        # paragraph.runs[0].font.size = Pt(10.5)
        paragraph.runs[0].font.color.rgb = RGBColor(235, 0, 0)


def main(old, new, *filenames, dryrun=False):
    """引数に対してreplace_textを実行する"""
    for filename in filenames:
        document = Document(filename)
        print("==filename:", filename, "==")
        # 本文書き換え
        for paragraph in document.paragraphs:
            replace_text(paragraph, old, new)
        # テーブル書き換え
        paragraphs = (paragraph for table in document.tables
                      for row in table.rows for cell in row.cells
                      for paragraph in cell.paragraphs)
        for paragraph in paragraphs:
            replace_text(paragraph, old, new)
        if not dryrun:
            document.save(filename)


def parse():
    """引数解析"""
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
    parser.add_argument('-V', '--version', action='store_true')
    return parser.parse_args()


if __name__ == '__main__':
    argv = parse()
    if argv.version:
        print('replacex:', VERSION)
        sys.exit(0)
    main(argv.old, argv.new, *argv.files, dryrun=argv.dryrun)
