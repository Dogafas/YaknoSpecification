# start.py
from utils import read_excel_and_return_data, create_table
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

class SpecGenerator:
    def __init__(self, excel_filepath="2055_ ЯКНО-ВВ(ВК)-6кВ_Соврудник.xlsm"):
        self.excel_filepath = excel_filepath
        self.data = None
        self.product_name = None
        self.headers = None

    def load_data(self):
        self.data = read_excel_and_return_data(self.excel_filepath)
        if self.data is None:
            print("Ошибка загрузки данных из Excel. Проверьте файл и путь.")
            return False
        return True

    def process_data(self):
        self.headers = self.extract_headers()
        self.product_name = self.extract_product_name()
        self.mechanical_locks = self.extract_mechanical_locks()
        if self.headers is None or self.product_name is None:
            print("Ошибка обработки данных. Проверьте структуру Excel файла.")
            return False
        return True


    def extract_headers(self):
        for item in self.data:
            if item.get("Структура") == "Структура" and item.get("Опция") == "Наименование":
                return {
                    "hdr_name": item.get("Опция", "Наименование"),
                    "hdr_unit": item.get("Примечание 4", "Ед. изм."),
                    "hdr_quantity": item.get("Примечание 1", "Кол-во"),
                }
        return {"hdr_name": "Наименование", "hdr_unit": "Ед. изм.", "hdr_quantity": "Кол-во"}

    def extract_product_name(self):
        for item in self.data:
            if item.get("Опция") == "Наименование изделия":
                return item.get("Значение")
        return "НУЖНО ВВЕСТИ НАЗВАНИЕ ПРОДУКТА"

    def extract_mechanical_locks(self):
        for item in self.data:
            if item.get("Скрыть строку, символы /*") == "/*":
                if item.get("Опция") and item.get("Опция").startswith("Механические блокировки:"):
                    return item
        return None

    def generate_document(self):
        if self.data is None:
            print("Данные не загружены. Выполните load_data()")
            return
        
        if not self.process_data():
            return

        doc = Document()
        heading = doc.add_heading(f'Техническая спецификация \n {self.product_name} \n (исполнение "Г")', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        create_table(doc, self.headers, self.data, self.mechanical_locks)
        doc.save(f'Техническая_спецификация_{self.product_name}.docx')
        print(f"Документ сохранен: Техническая_спецификация_{self.product_name}.docx")

if __name__ == "__main__":
    generator = SpecGenerator()
    if generator.load_data():
        generator.generate_document()