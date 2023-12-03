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

    def read_row(
        self,
        row: int = None,
        column_from: int = None,
        column_to: int = None,
    ):
        if row is None:
            row = self.active_row
        column_from = column_from if column_from is not None else 1
        column = column_from
        column_num = column_to - column + 1 if column_to is not None else None

        contents = []
        content = self.read_from_cells(row=row, column=column)
        while (
            column_num is None
            and content is not None  # read until None
            or column_num is not None
            and column - column_from < column_num  # read until reaching column_num
        ):
            contents.append(content)
            column += 1
            content = self.read_from_cells(row=row, column=column)
        return contents

    def read_column(
        self,
        column: int = None,
        row_from: int = None,
        row_to: int = None,
    ):
        if column is None:
            column = self.active_column
        row_from = row_from if row_from is not None else 1
        row = row_from
        row_num = row_to - row + 1 if row_to is not None else None

        contents = []
        content = self.read_from_cells(row=row, column=column)
        while (
            row_num is None
            and content is not None  # read until None
            or row_num is not None
            and row - row_from < row_num  # read until reaching column_num
        ):
            contents.append(content)
            row += 1
            content = self.read_from_cells(row=row, column=column)
        return contents

    def read_area(
        self,
        row_from: int = None,
        row_to: int = None,
        column_from: int = None,
        column_to: int = None,
        header: bool = False,
    ):
        row_from = row_from if row_from is not None else 1
        column_from = column_from if column_from is not None else 1

        if header:
            column_names = self.read_row(
                row=row_from, column_from=column_from, column_to=column_to
            )
            row_from += 1

        row_contents = []
        for row in range(row_from, row_to + 1):
            row_content = self.read_row(
                row=row, column_from=column_from, column_to=column_to
            )
            if header:
                row_dict = {}
                for i in range(0, len(column_names)):
                    row_dict[column_names[i]] = row_content[i]
                row_contents.append(row_dict)
            else:
                row_contents.append(row_content)

        return row_contents
