import argparse
import io
import re
from urllib import request

from openpyxl import load_workbook

from xcearch.htmlparser import HTMLLinkParser


def main():
    parser = argparse.ArgumentParser(description='Search in all excel documents on one or more websites.')
    parser.add_argument('text', metavar='T', type=str,
                        help='The text to search for in the excel documents')
    parser.add_argument('urls', metavar='URL', type=str, nargs='+',
                        help='The website linking to excel documents')
    args = parser.parse_args()

    sites = []
    for url in args.urls:
        sites.append((url, HTMLLinkParser(url).documents))

    pattern = re.compile(args.text, re.IGNORECASE)

    for s in sites:
        for url in s[1]:
            print(url)
            f = io.BytesIO(request.urlopen(url[1]).read())
            wb = load_workbook(filename=f, read_only=True)
            result = []
            for ws in wb:
                for row in ws.rows:
                    for cell in row:
                        try:
                            if pattern.search(cell.value):
                                result.append((ws.title, cell.row, cell.column))
                        except TypeError:
                            continue

            for (ws, row, column) in result:
                print(ws, row, column)


if __name__ == "__main__":
    main()
