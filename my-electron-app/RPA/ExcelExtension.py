from RPA.Excel.Application import Application as ExcelApplication


def index_num_to_str(index: int):
    result = ""
    while index != 0:
        result = chr(ord("A") + index % 26 - 1) + result
        index = index // 26
    return result


def index_str_to_num(character: str):
    result = 0
    length = len(character)
    base = 1
    for i in range(0, length):
        result += (ord(character[length - i - 1]) - ord("A") + 1) * base
        base *= 26
    return result


class ExcelApplicationExtension(ExcelApplication):
    def __init__(self):
        super().__init__()
        self.active_row = 1
        self.active_column = "A"
        self.cached_header_row_value= None
        self.cached_header_row_index = -1
        
    def fetch_header_row_value(self, header_index: int):
        if self.cached_header_row_index != header_index:
            self.cached_header_row_index = header_index
            self.cached_header_row_value = self.worksheet.Rows(header_index).Value[0]
        return self.cached_header_row_value
    
    def set_active_worksheet(
        self, sheetname: str = None, sheetnumber: int = None
    ):
        self.cached_header_row_index = -1
        if sheetnumber:
            self.worksheet = self.workbook.Worksheets(int(sheetnumber))
        elif sheetname:
            self.worksheet = self.workbook.Worksheets(sheetname)
            
    def read_row(
        self,
        row: int = None,
        column_from: str = None,
        column_to: str = None,
    ):
        rangeStr = column_from + str(row) + ":" + column_to + str(row)
        contents = [None] + list(self.worksheet.Range(rangeStr).Value[0])
        return contents

    def read_row_with_header(
        self,
        row: int = None,
        column_from: str = None,
        column_to: str = None,
        header_row: int = None,
    ):
        rangeStr = column_from + str(row) + ":" + column_to + str(row)
        contents = list(self.worksheet.Range(rangeStr).Value[0])
        index_from = index_str_to_num(column_from) - 1
        index_to = index_str_to_num(column_to)
        headers = self.fetch_header_row_value(header_row)[index_from:index_to]
        contents_dict = {}
        for header, content in zip(headers, contents):
            contents_dict[header] = content
        return contents_dict

    def write_row(
        self,
        row: int = None,
        row_content=None,
        column_from: str = None,
        column_to: str = None,
    ):
        if row == self.cached_header_row_index:
            self.cached_header_row_index = -1
        
        row_value = row_content[1:]
        rangeStr = column_from + str(row) + ":" + column_to + str(row)
        self.worksheet.Range(rangeStr).Value = row_value

    def write_row_with_header(
        self,
        row: int = None,
        row_content=None,
        column_from: str = None,
        column_to: str = None,
        header_row: int = None,
    ):
        if row == self.cached_header_row_index:
            self.cached_header_row_index = -1
        
        index_from = index_str_to_num(column_from) - 1
        index_to = index_str_to_num(column_to)
        headers = self.fetch_header_row_value(header_row)[index_from:index_to]
        
        row_value = []
        for header in headers:
            if header in row_content.keys():
                row_value.append(row_content[header])
            else:
                row_value.append("")

        rangeStr = column_from + str(row) + ":" + column_to + str(row)
        self.worksheet.Range(rangeStr).Value = row_value

    def write_column(
        self,
        column: str = None,
        column_content=None,
        row_from: int = None,
        row_to: int = None,
    ):
        if row_from <= self.cached_header_row_index and row_to >= self.cached_header_row_index:
            self.cached_header_row_index = -1
        
        column_value = [(content,) for content in column_content[1:]]
        rangeStr = column + str(row_from) + ":" + column + str(row_to)
        self.worksheet.Range(rangeStr).Value = column_value

    def read_column(
        self,
        column: str = None,
        row_from: int = None,
        row_to: int = None,
    ):
        rangeStr = column + str(row_from) + ":" + column + str(row_to)
        contents = [None] + [content[0] for content in self.worksheet.Range(rangeStr).Value]
        return contents

    def read_area(
        self,
        row_from: int = None,
        row_to: int = None,
        column_from: str = None,
        column_to: str = None,
    ):
        row_contents = [None]
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
        header_row: int = None
    ):
        index_from = index_str_to_num(column_from) - 1
        index_to = index_str_to_num(column_to)
        headers = self.fetch_header_row_value(header_row)[index_from:index_to]

        row_contents = [None]
        for row in range(row_from, row_to + 1):
            row_content = self.read_row(
                row=row, column_from=column_from, column_to=column_to
            )
            row_dict = {}
            for i in range(0, len(headers)):
                row_dict[headers[i]] = row_content[i]
            row_contents.append(row_dict)
        return row_contents

    def data_type_to_text(
        self,
        row_from: int = None,
        row_to: int = None,
        column_from: str = None,
        column_to: str = None,
    ):
        rangeStr = column_from + str(row_from) + ":" + column_to + str(row_to)
        self.worksheet.Range(rangeStr).NumberFormat = "@"

    class WorkbookDict:
        def __init__(self):
            self.workbook_contents = {}
            self.headers = None
            self.header_row = None
            self.text_columns = []

        def contains_name(self, name: str):
            return name in self.workbook_contents.keys()

        def names(self):
            return [None] + self.workbook_contents.keys()

        def column_data_type_to_text(self, column: str):
            self.text_columns.append(column)

        def add_workbook(self, name: str):
            self.workbook_contents[name] = []

        def set_headers(self, headers: list[str], header_row: int):
            self.headers = headers
            self.header_row = header_row

        def add_row(self, name: str, row_content=None):
            row_value = row_content
            if not self.contains_name(name):
                self.add_workbook(name)
            self.workbook_contents[name].append(row_value)

        def generate_workbook_files(self, path=None):
            for name, row_contents in self.workbook_contents.items():
                app = ExcelApplicationExtension()
                app.open_application(visible=True)
                app.add_new_workbook()
                app.add_new_sheet(name)

                if self.headers is None:
                    for column in self.text_columns:
                        app.data_type_to_text(
                            row_from=1,
                            row_to=len(row_contents),
                            column_from=column,
                            column_to=column,
                        )
                    i = 1
                    for row_content in row_contents:
                        app.write_row(
                            row=i,
                            row_content=row_content,
                            column_from="A",
                            column_to=index_num_to_str(len(row_content)),
                        )
                        i += 1
                else:
                    app.write_row(
                        row=self.header_row,
                        row_content=self.headers,
                        column_from="A",
                        column_to=index_num_to_str(len(self.headers) - 1),
                    )
                    for column in self.text_columns:
                        app.data_type_to_text(
                            row_from=self.header_row + 1,
                            row_to=self.header_row + len(row_contents),
                            column_from=column,
                            column_to=column,
                        )
                    for i in range(len(row_contents)):
                        app.write_row_with_header(
                            row=self.header_row + 1 + i,
                            row_content=row_contents[i],
                            column_from="A",
                            column_to=index_num_to_str(len(self.headers) - 1),
                            header_row=self.header_row
                        )

                app.save_excel_as(filename=path + name + ".xls", file_format=56)
                app.close_document()