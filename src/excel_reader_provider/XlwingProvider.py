import xlwings as xw
from xlwings import App, Book

from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider


class XlwingProvider(ExcelReaderProvider):

    def __init__(self):
        self._app: App = xw.App(visible=False)
        self.name_to_workbook: dict[str, object] = {}

    def get_workbook(self, path: str):
        # Open an existing workbook
        if self.name_to_workbook.get(path):
            return self.name_to_workbook.get(path)

        wb = self._app.books.open(path)
        self.name_to_workbook[path] = wb
        return wb

    def get_worksheet(self, workbook, sheet_name: str):
        ws = workbook.sheets[sheet_name]
        return ws

    def change_value_at(self, worksheet, row, column, value):
        worksheet.range(row, column).value = value
        return True

    def get_value_at(self, worksheet, row, column):
        return worksheet.range((row, column)).value

    def delete_contents(self, worksheet, start_cell, end_cell):
        worksheet.range(start_cell + ":" + end_cell).clear_contents()
        return True

    def save(self, workbook):
        workbook.save()

    def close(self, workbook: Book):
        path_to_workbook: str = workbook.fullname
        workbook.close()
        if self.name_to_workbook.get(path_to_workbook) is not None:
            del self.name_to_workbook[path_to_workbook]

    def quit_session(self):
        self._app.quit()
