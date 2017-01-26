from html.parser import HTMLParser
from urllib import request
from urllib.parse import urlparse
from urllib.parse import urlunparse


class HTMLLinkParser(HTMLParser):
    link = ""
    documents = []

    def __init__(self, url):
        super().__init__()
        self.base = urlparse(url)
        html = request.urlopen(url).read().decode('utf-8')
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
