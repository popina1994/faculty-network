import openpyxl as px
import re
from openpyxl import Workbook
class GraphXlsxConverter:
    SOURCE_COLUMN_IDX = 1
    TARGET_COLUMN_IDX = 2
    COLUMN_NAMES_ROW_ID = 1

    COLUMN_NAME_AUTHORS_FULL_NAME = "konkatenirano"
    COLUMN_NAME_OF_AUTHORS_OPT_1 = "autori"
    COLUMN_NAME_OF_AUTHORS_OPT_2 = "autori2"

    COLUMN_NAME_OF_YEAR = "godina"

    COLUMN_NAME_SOURCE_NODE = "Source"
    COLUMN_NAME_DEST_NODE = "Target"

    authors = []

    def find_authors_full_name_column_idx(sheet):
        if (sheet.cell(column = 1, row = GraphXlsxConverter.COLUMN_NAMES_ROW_ID) is None):
            return 0
        for col_idx in range(1, sheet.max_column):
            cell = sheet.cell(column = col_idx, row = GraphXlsxConverter.COLUMN_NAMES_ROW_ID)
            if (cell.value == GraphXlsxConverter.COLUMN_NAME_AUTHORS_FULL_NAME):
                return col_idx
        return 0

    def init(work_book_authors_name):
        work_book_read = px.load_workbook(filename=work_book_authors_name)
        sheet_names = work_book_read.get_sheet_names();
        for sheet_name in sheet_names:
            sheet = work_book_read[sheet_name]
            authors_full_name_col_idx = GraphXlsxConverter.find_authors_full_name_column_idx(sheet)
            if (authors_full_name_col_idx == 0):
                continue
            
            for row_idx in range(GraphXlsxConverter.COLUMN_NAMES_ROW_ID + 1, sheet.max_row):
                cell_value = sheet.cell(column = authors_full_name_col_idx, row = row_idx).value
                if ((cell_value == "") or (cell_value is None)):
                    continue
                GraphXlsxConverter.authors.append(cell_value.strip())

    def __init__(self, work_book_read_name):
        self.work_book_read_name = work_book_read_name

    def find_autors_column_idx(sheet):
        if (sheet.cell(column = 1, row = GraphXlsxConverter.COLUMN_NAMES_ROW_ID) is None):
            return "", 0
        for col_idx in range(1, sheet.max_column):
            cell = sheet.cell(column = col_idx, row = GraphXlsxConverter.COLUMN_NAMES_ROW_ID)
            if (cell.value == GraphXlsxConverter.COLUMN_NAME_OF_AUTHORS_OPT_1):
                return GraphXlsxConverter.COLUMN_NAME_OF_AUTHORS_OPT_1, col_idx
            if (cell.value == GraphXlsxConverter.COLUMN_NAME_OF_AUTHORS_OPT_2):
                return GraphXlsxConverter.COLUMN_NAME_OF_AUTHORS_OPT_2, col_idx
        return "", 0

    def find_year_column_idx(sheet):
        if (sheet.cell(column = 1, row = GraphXlsxConverter.COLUMN_NAMES_ROW_ID) is None):
            return "", 0
        for col_idx in range(1, sheet.max_column):
            cell = sheet.cell(column = col_idx, row = GraphXlsxConverter.COLUMN_NAMES_ROW_ID)
            if (cell.value == GraphXlsxConverter.COLUMN_NAME_OF_YEAR):
                return GraphXlsxConverter.COLUMN_NAME_OF_YEAR, col_idx

        return "", 0


    def init_sheet_names(sheet):
        sheet.cell(row = GraphXlsxConverter.COLUMN_NAMES_ROW_ID, 
                   column = GraphXlsxConverter.SOURCE_COLUMN_IDX).value = GraphXlsxConverter.COLUMN_NAME_SOURCE_NODE
        sheet.cell(row = GraphXlsxConverter.COLUMN_NAMES_ROW_ID, 
                   column = GraphXlsxConverter.TARGET_COLUMN_IDX).value = GraphXlsxConverter.COLUMN_NAME_DEST_NODE

    def generate_edges_of_authors(self, sheet, authors):
        for author_1 in authors:
                for author_2 in authors:
                        sheet.cell(row = self.row_write_idx, column = GraphXlsxConverter.SOURCE_COLUMN_IDX).value = author_1
                        sheet.cell(row = self.row_write_idx, column = GraphXlsxConverter.TARGET_COLUMN_IDX).value = author_2
                        self.row_write_idx += 1

    def convert_xlsx_to_graph(self):
        work_book_read = px.load_workbook(filename=self.work_book_read_name)
        work_book_write = Workbook()
        sheet_write_active = work_book_write.active
        sheet_names = work_book_read.get_sheet_names();
        for sheet_name in sheet_names:
            sheet = work_book_read[sheet_name]
            option, authors_col_idx = GraphXlsxConverter.find_autors_column_idx(sheet)
            _, year_col_idx = GraphXlsxConverter.find_year_column_idx(sheet)
            if (authors_col_idx == 0):
                continue
            
            GraphXlsxConverter.init_sheet_names(sheet_write_active)
            self.row_write_idx = GraphXlsxConverter.COLUMN_NAMES_ROW_ID + 1
            for row_idx in range(GraphXlsxConverter.COLUMN_NAMES_ROW_ID + 1, sheet.max_row):
                year_value = sheet.cell(column = year_col_idx, row = row_idx).value
                if (type(year_value) is int):
                    year_value = int(year_value)
                else:
                    if (type(year_value) is type(None)):
                        print("None")
                    else:
                        print("Irregular" + year_value + "\n")
                    continue
                if ((year_value < 2000) or (year_value > 2016)):
                    continue
                cell_value = sheet.cell(column = authors_col_idx, row = row_idx).value
                if ((cell_value == "") or (cell_value is None)):
                    continue
                
                if (option == GraphXlsxConverter.COLUMN_NAME_OF_AUTHORS_OPT_2):
                    # Remove the first {, and at the end }"
                    cell_value =  "\"" + cell_value[1:-2];

                list_of_authors_not_formated = cell_value.split(",")
                list_of_authors = []
                for author_not_format in list_of_authors_not_formated:
                    author = author_not_format.strip()
                    author = re.sub(' +',' ',author);
                    
                    if (option == GraphXlsxConverter.COLUMN_NAME_OF_AUTHORS_OPT_2):
                        # remove ""
                        author = author[1 : -1]
                    if (author.find(".") != -1):
                        continue
                    if (author not in GraphXlsxConverter.authors):
                        continue
                    list_of_authors.append(author)
                self.generate_edges_of_authors(sheet_write_active, list_of_authors)

        work_book_write.save("GRAPH_" + self.work_book_read_name)
        work_book_read.close()


