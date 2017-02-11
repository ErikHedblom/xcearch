import argparse
import io
import re
from html.parser import HTMLParser
from urllib.parse import urlparse, urlunparse
from urllib.request import Request, urlopen

from openpyxl import load_workbook


class HTMLLinkParser(HTMLParser):
    link = ""
    documents = []

    def __init__(self, url):
        super().__init__()
        self.base = urlparse(url)
        req = Request(url, headers={'User-Agent': 'Excearch/0.1'})
        html = urlopen(req).read().decode('utf-8')
        self.feed(html)

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            for name, value in attrs:
                if name == "href" and value.endswith('.xlsx'):
                    u = urlparse(value)
                    if not u[0]:
                        self.link = urlunparse((self.base[0], self.base[1], u[2], u[3], u[4], u[5]))
                    else:
                        self.link = urlunparse(u)

    def handle_data(self, data):
        if self.link:
            self.documents.append((data, self.link))
            self.link = ""


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
            f = io.BytesIO(urlopen(url[1]).read())
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
