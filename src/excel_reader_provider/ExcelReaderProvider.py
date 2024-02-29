from abc import abstractmethod


class ExcelReaderProvider:
    @abstractmethod
    def get_workbook(self, path: str):
        pass

    @abstractmethod
    def get_worksheet(self, workbook, sheet_name: str):
        pass

    @abstractmethod
    def change_value_at(self, worksheet, row, column, value):
        pass

    @abstractmethod
    def get_value_at(self, worksheet, row, column):
        pass

    @abstractmethod
    def save(self, workbook):
        pass

    @abstractmethod
    def close(self, workbook):
        pass

    @abstractmethod
    def quit_session(self):
        pass
