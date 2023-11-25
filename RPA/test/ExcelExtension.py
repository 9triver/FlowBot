from RPA.Excel.Application import Application as ExcelApplication


class ExcelApplicationExtension(ExcelApplication):
    def __init__(self):
        super().__init__()
        self.active_row = 1
        self.active_column = 1

    def move_active_cell(self, row_change: int = 0, column_change: int = 0):
        self.active_row += row_change
        if self.active_row < 1:
            self.active_row = 1
        self.active_column += column_change
        if self.active_column < 1:
            self.active_column = 1

    def set_active_cell(self, row: int = 0, column: int = 0):
        self.active_row = row
        self.active_column = column

    def read_row(self, row: int = None, header: bool = False, column_num: int = None):
        if row is None:
            row = self.active_row

        if header:
            column_names = self.read_row(1)
            contents = {}
            for column in range(1, len(column_names) + 1):
                contents[column_names[column - 1]] = self.read_from_cells(
                    row=row, column=column
                )
            return contents

        contents = []
        column = 1
        content = self.read_from_cells(row=row, column=column)
        while (
            column_num is None
            and content is not None
            or column_num is not None
            and column <= column_num
        ):
            contents.append(content)
            column += 1
            content = self.read_from_cells(row=row, column=column)
        return contents
