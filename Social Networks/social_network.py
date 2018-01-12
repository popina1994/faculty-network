import language.xlsx_language_converter as lc
import graph.graph_xlsx_converter as gc

def latin_to_cyrilic():
    work_book_names = ["ugovori.xlsx", "casopisi.xlsx", "doma_konferencije.xlsx", "medj_konferencije.xlsx"]
    for work_book_name in work_book_names:
        xlsx_converter = lc.XlsxCyrilicLatinConverter(work_book_name)
        xlsx_converter.convert_cyrilic_to_latin()

def xlsx_to_graph():
     work_book_names = ["LATINmedj_konferencije.xlsx", "LATINdoma_konferencije.xlsx", "LATINcasopisi.xlsx"]
     gc.GraphXlsxConverter.init("LATINugovori.xlsx")
     for work_book_name in work_book_names:
        graph_converter = gc.GraphXlsxConverter(work_book_name)
        graph_converter.convert_xlsx_to_graph()



xlsx_to_graph()
    
