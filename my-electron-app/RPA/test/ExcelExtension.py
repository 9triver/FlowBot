from RPA.Excel.Application import Application as ExcelApplication


def index_to_character(index: int):
    result = ''
    while index != 0:
        result = chr(ord('A') + index % 26 - 1) + result
        index = index // 26
    return result

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
        column_to: int = None
    ):
        index_from = column_from - 1
        index_to = column_to
        contents = self.worksheet.Rows(row).Value[0][index_from : index_to]
        return contents

    def read_row_with_header(
        self,
        row: int = None,
        column_from: int = None,
        column_to: int = None,
        header_row: int = None,
    ):
        index_from = column_from - 1
        index_to = column_to
        contents = self.worksheet.Rows(row).Value[0][index_from : index_to]
        headers = self.worksheet.Rows(header_row).Value[0][index_from : index_to]
        contents_dict = {}
        for header, content in zip(headers, contents):
            contents_dict[header] = content
        return contents_dict
        
    def write_row(self, row: int = None, 
                   row_content=None,
                   column_from: int = None,
                   column_to: int = None):
        row_value = row_content
        rangeStr = str(index_to_character(column_from)) + str(row) + ':' + str(index_to_character(column_to)) + str(row)
        self.worksheet.Range(rangeStr).Value = row_value

    def write_row_with_header(self, 
                               row: int = None, 
                               row_content = None, 
                               column_from: int = None,
                               column_to: int = None, 
                               header_row: int = None):
        headers = self.read_row(row=header_row, column_from=column_from, column_to=column_to)
        row_value = []
        for header in headers:
            if header in row_content.keys():
                row_value.append(row_content[header])
            else:
                row_value.append('')
        rangeStr = str(index_to_character(column_from)) + str(row) + ':' + str(index_to_character(column_to)) + str(row)
        self.worksheet.Range(rangeStr).Value = row_value

    def write_column(self, column: int = None, column_content=None, row_from: int = None, row_to: int = None):
        value = [(content,) for content in column_content]
        rangeStr = str(index_to_character(column)) + str(row_from) + ':' + str(index_to_character(column)) + str(row_to)
        self.worksheet.Range(rangeStr).Value = value

    def read_column(
        self,
        column: int = None,
        row_from: int = None,
        row_to: int = None,
    ):
        if column is None:
            column = self.active_column

        index_from = row_from - 1
        index_to = row_to

        contents = [content[0] for content in self.worksheet.Columns(column).Value[index_from : index_to]]
        return contents

    def read_area(
        self,
        row_from: int = None,
        row_to: int = None,
        column_from: int = None,
        column_to: int = None,
    ):
        row_contents = []
        for row in range(row_from, row_to + 1):
            row_content = self.read_row(
                row=row, column_from=column_from, column_to=column_to
            )
            row_contents.append(row_content)

        return row_contents

    def read_area_with_header(
        self,
        row_from: int = None,
        row_to: int = None,
        column_from: int = None,
        column_to: int = None,
    ):
        headers = self.read_row(
            row=row_from, column_from=column_from, column_to=column_to
        )
        row_from += 1

        row_contents = []
        for row in range(row_from, row_to + 1):
            row_content = self.read_row(
                row=row, column_from=column_from, column_to=column_to
            )
            row_dict = {}
            for i in range(0, len(headers)):
                row_dict[headers[i]] = row_content[i]
            row_contents.append(row_dict)
        return row_contents

    class WorkbookDict:
        def __init__(self):
            self.workbook_contents = {}
            self.headers = None
            self.header_row = None

        def contains_name(self, name: str):
            return name in self.workbook_contents

        def add_workbook(self, name: str):
            self.workbook_contents[name] = []

        def set_headers(self, headers: list[str], header_row: int):
            self.headers = headers
            self.header_row = header_row

        def add_row(self, name: str, row_content=None):
            if not self.contains_name(name):
                self.add_workbook(name)
            self.workbook_contents[name].append(row_content)

        def generate_workbook_files(self, path="./"):
            for name, row_contents in self.workbook_contents.items():
                app = ExcelApplicationExtension()
                app.open_application(visible=True)
                app.add_new_workbook()
                app.add_new_sheet(name)

                if self.headers is None:
                    i = 1
                    for row_content in row_contents:
                        app.write_row(row=i, row_content=row_content)
                        i += 1
                else:
                    app.write_row(row=self.header_row, row_content=self.headers)
                    i = 1
                    for row_content in row_contents:
                        if i == self.header_row:
                            i += 1
                        row_value = []
                        for header in self.headers:
                            if header in row_content.keys():
                                row_value.append(row_content[header])
                        app.write_row(
                            row=i, row_content=row_value
                        )
                        i += 1

                app.save_excel_as(filename=path + name + ".xls", file_format=56)
                app.close_document()
