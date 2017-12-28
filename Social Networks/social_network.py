import numpy as np
import openpyxl as px
import language_converter as lc

work_book = px.load_workbook(filename='ugovori.xlsx', data_only=True)
sheet_ranges = work_book['Zaposleni']
for row in sheet_ranges.iter_rows(min_row=1, max_col = 3, max_row=2):
    for cell in row:
        print(lc.CyrilicToLatin.convertCyrilicToLatin(cell.internal_value))