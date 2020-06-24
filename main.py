import collections
from datetime import datetime

import pandas
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

from config import SINCE_YEAR, TEMPLATE, XLS_DATA_FILE, HOST_PORT


def parse_wines_from_xls(file_name):
    wine_data = collections.defaultdict(list)
    wines_df = pandas.read_excel(file_name)

    wines_df = wines_df.fillna('')

    for wine in wines_df.values:
        wine_data[wine[0]].append({
            'category': wine[0],
            'name': wine[1],
            'grade': wine[2],
            'price': wine[3],
            'picture': wine[4],
            'promo': wine[5]
        })
    return sorted(wine_data.items(), key=lambda item: item)


def calc_years_from():
    return datetime.now().year - SINCE_YEAR


def prepare_template():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    return env.get_template(TEMPLATE)


def main():
    template = prepare_template()

    rendered_page = template.render(
        years=calc_years_from(),
        wines=parse_wines_from_xls(XLS_DATA_FILE)
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(HOST_PORT, SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
