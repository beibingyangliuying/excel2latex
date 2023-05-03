import xlwings as xlw


class ExcelReader:
    def __init__(self, file_path: str, sheet_name: str, scope: str):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.scope = scope

    def __enter__(self):
        try:
            self._book: xlw.Book = xlw.Book(self.file_path, read_only=True)
        except FileNotFoundError as e:
            raise e

        try:
            self._sheet: xlw.Sheet = self._book.sheets[self.sheet_name]
        except IndexError as e:
            raise e

        # TODO：可以考虑判断scope的合法性
        self.range: xlw.Range = self._sheet[self.scope]
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        app = self._book.app
        self._book.close()
        app.quit()
