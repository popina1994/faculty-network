import xlsx_language_converter as lc

def run():
    work_book_names = ["ugovori.xlsx", "casopisi.xlsx", "doma_konferencije.xlsx", "medj_konferencije.xlsx"];
    for work_book_name in work_book_names:
        xlsx_converter = lc.xlsx_cyrilic_latin_converter(work_book_name)
        xlsx_converter.convert_cyrilic_to_latin()

run()