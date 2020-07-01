from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
import math

from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
import xlrd

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

excel_data = pandas.read_excel('wine3.xlsx', sheet_name='wine')

dict_wins = defaultdict(list)
columns = excel_data.columns.ravel()
for i, lines in excel_data.iterrows():
    dict = {}
    for column in columns:
        dict[column] = str(lines[column])
    dict_wins[lines['category']].append(dict)

template = env.get_template('template.html')

rendered_page = template.render(
    time_title=datetime.datetime.now().year - 1920,
    wins=dict_wins,
    categorylist=list(dict_wins.keys()),
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
