import re
from openpyxl import load_workbook

pattern = re.compile('susd4', re.IGNORECASE)

wb = load_workbook(filename='test.xlsx', read_only=True)
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
