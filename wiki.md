FlowBot Blocks是面向FlowBot所支持的应用场景（例如Excel自动操作）在Blockly中的自定义[Block](https://developers.google.com/blockly/guides/create-custom-blocks/overview)。本项目中包括：

**Script Header**

```
from robocorp.tasks import task

from ExcelExtension import ExcelApplicationExtension as ExcelApplication

@task
def solve_challenge():
```

**Workbook operation**

* Open Workbook {path} As {var} : 打开Excel文档
    * path: Excel 文档路径
    
    * var: 表示文档的变量名
    
    * ```
        {var} = ExcelApplication()
        {var}.open_application(visible=True)
        {var}.open_workbook({path})
        ```
    
* Add Workbook As {var} : 新建Excel文档

    * var: 表示文档的变量名

    * ```
        {var} = ExcelApplication()
        {var}.open_application(visible=True)
        {var}.add_new_workbook()
        ```

* {workbook} Save Workbook {path} : 保存 Excel 文档
  
    * workbook: Excel 文档变量名
    
    * path: 目标保存路径，为空表示在文档原位置覆盖保存
    
    * ```
        # path is empty
        {workbook}.save_excel()
        # path is not empty
        {workbook}.save_excel_as(filename={path})
        ```
    

**Sheet operation**

* {workbook} Create Sheet {name} : 新建表

    * workbook: Excel 文档变量名

    * name: 新建的表名

    * ```
        {workbook}.add_new_sheet({name})
        ```

* {workbook} Set Active Sheet {name} : 改变活跃中的表

    * workbook: Excel 文档变量名

    * name: 将要设为活跃的表名

    * ```
        {workbook}.set_active_worksheet({name})
        ```

* merge sheet (合并表)

**Row and Column and Cell operation**

* {workbook} Move Active Cell {row_change} {column_change} : 移动活跃单元格

    * workbook: Excel 文档变量名

    * row_change: 行变化，默认为0

    * column_change: 列变化，默认为0

    * ```
        {workbook}.move_active_cell(row_change={row_change}, column_change={column_change})
        ```

* {workbook} Set Active Cell {row} {column} : 设置活跃单元格

    * workbook: Excel 文档变量名

    * row: 行号

    * column: 列号

    * ```
        {workbook}.set_active_cell(row={row}, column={column})
        ```

* {workbook} Fetch Cell  {row} {column} As {var} : 获取单元格

    * workbook: Excel 文档变量名

    * row: 行号，为空则采用当前获取行

    * column: 列号，为空则采用当前活跃列

    * var: 表示获取结果的变量名

    * ```
        {var} = {workbook}.read_from_cells(row={row}, column={column})
        {var} = str({var}) if {var} is not None else None
        ```

* {workbook} Fetch Row {row} {column_from} {column_to} {header_row} As {var} :  获取一行

    * workbook: Excel 文档变量名

    * row: 行号，为空则采用当前活跃行

    * column_from: 起点列号，为空则采用第一列

    * column_to: 终点列号，为空则读取到空值为止

    * var: 表示获取结果的变量名

    * ```
        {var} = {workbook}.read_row(row={row}, column_from={column_from}, column_to={column_to})
        ```
    
* {workbook} Fetch Row {row} {column_from} {column_to}  With Header {header_row} As {var} :  获取的一行，以某行作为列名
  
    * workbook: Excel 文档变量名
    
    * row: 行号，为空则采用当前活跃行
    
    * column_from: 起点列号，为空则采用第一列
    
    * column_to: 终点列号，为空则读取到空值为止
    
    * header_row: header 所在行号
    
    * var: 表示获取结果的变量名
    
    * ```
        {var} = {workbook}.read_row_with_header(row={row}, column_from={column_from}, column_to={column_to}, header_row={header_row})
        ```
    
* {workbook} Fetch Column {column} {row_from} {row_to} As {var} :  获取一列
  
    * workbook: Excel 文档变量名
    
    * column: 列号，为空则采用当前活跃列
    
    * row_from: 起点行号
    
    * row_to: 终点行号
    
    * var: 表示获取结果的变量名
    
    * ```
        {var} = {workbook}.read_column(column={column}, row_from={row_from}, row_to={row_to})
        ```
    
* {workbook} Fetch area {row_from} {row_to} {column_from} {column_to} {with_header} As {var} : 获取一个区域

    * workbook: Excel 文档变量名

    * row_from: 起点行号

    * row_to: 终点行号

    * column_from: 起点列号

    * column_to: 终点列号

    * var: 表示获取结果的变量名

    * ```
        {var} = {workbook}.read_area(row_from={row_from}, row_to={row_to}, column_from={column_from}, column_to={column_to})
        ```
    
* {workbook} Fetch area with header {row_from} {row_to} {column_from} {column_to} {with_header} As {var} : 获取一个区域，以第一行作为列名

    * workbook: Excel 文档变量名

    * row_from: 起点行号

    * row_to: 终点行号

    * column_from: 起点列号

    * column_to: 终点列号

    * var: 表示获取结果的变量名

    * ```
        {var} = {workbook}.read_area_with_header(row_from={row_from}, row_to={row_to}, column_from={column_from}, column_to={column_to})
        ```

* {workbook} Write Row {row} {row_content} {header_row} :  写入行

    * workbook: Excel 文档变量名

    * row_content: 待写入的行

    * column_from: 起点列号

    * column_to: 终点列号

    * ```
        {workbook}.write_row(row={row}, row_content={row_content}, column_from={column_from}, column_to={column_to})
        ```

* {workbook} Write Row {row} {row_content} {column_from} {column_to} With Header {header_row} :  写入带有列名信息的行

    * workbook: Excel 文档变量名

    * row_content: 待写入的行

    * column_from: 起点列号

    * column_to: 终点列号

    * header_row: header 所在行号

    * ```
        {workbook}.write_row_with_header(row={row}, row_content={row_content}, column_from={column_from}, column_to={column_to}, header_row={header_row})
        ```


* {workbook} Write Column {column} {column_content} : 写入列

    * workbook: Excel 文档变量名

    * column_content: 待写入的列

    * row_from: 起点行号

    * row_to: 终点行号

    * ```
        {workbook}.write_column(column={column}, column_content={column_content}, row_from={row_from}, row_to={row_to})
        ```

* {workbook} Set Cell {row} {column} {value}: 设置单元格的值

    * workbook: Excel 文档变量名

    * row: 行号，为空则采用当前获取行

    * column: 列号，为空则采用当前活跃列

    * value: 待写入的值

    * ```
        {workbook}.write_to_cells(row={row}, column={column}, value={value}, number_format='@')
        ```

**Workbook Dictionary operation**

- Make workbook dictionary {var}: 生成一个workbook集合

    - var: 生成的workbook集合变量名

    - ```
        {var} = ExcelApplication.WorkbookDict()
        ```

- {workbook_dict} set headers {headers} {header_row}: 设置workbook集合的表头，设置后将会按照对应表头插入内容

    - workbook_dict: workbook集合

    - headers: 设置的表头

    - header_row: headers所在行号

    - ```
        {workbook_dict}.set_headers(headers={headers}, header_row={header_row})
        ```

- {workbook_dict} add row {name} {row_content}: 向一个workbook新增一行

    - workbook_dict: workbook集合

    - name: 需要新增一行的workbook名

    - row_content: 新增内容

    - ```
        {workbook_dict}.add_row(name={name}, row_content={row_content})
        ```

- {workbook_dict} generate workbook files {path}: 生成excel文件

    - workbook_dict: workbook集合

    - path: 生成目录，默认为 './'

    - ```
        {workbook_dict}.generate_workbook_files(path={path})
        ```

**Variable operation**

- Set {var} to {exp} : 赋值操作

    - var:  表示待赋值变量的变量名 

    - exp: 表示值的表达式，可以包含简单运算

    - ```
        {var} = {exp}
        ```

        

**Control block**

- if {condition} : 分支控制块 if

    - condition: 条件块，详见 Condition block 部分

    - ```
        if {condition}:
        ```

    - 块内缩进+1

- else if {condition} : 分支控制块 else if

    - condition: 条件块，详见 Condition block 部分

    - ```
        elif {condition}:
        ```

    - 块内缩进+1

- else : 分支控制块 else

    - ```
        else:
        ```

    - 块内缩进+1

- for {var} from {start} to {end} : 循环控制块，带int型循环变量，前闭后闭，start <= end

    - var: 循环变量

    - start: var 最小值

    - end: var 最大值

    - ```
        for {var} in range({start}, {end} + 1):
        ```

    - 块内缩进+1

- for each {var} in {iterable_var} : 循环控制块，for each 循环，针对可迭代容器

    - var: 表示容器中每个元素的变量名

    - iterable_var: 表示一个可迭代容器

    - ```
        for {var} in {iterable_var}:
        ```

    - 块内缩进+1



**Condition block**

- {value_type} {exp1} {comparation} {exp2} : 比较条件块

    - value_type: 表达式数据类型，三种可选项为 int float str

    - exp1, exp2: 参与比较的两个表达式

    - comparation: 比较运算符，五种可选项为 < <= == >= >

    - ```
        {value_type}({exp1}) {comparation} {value_type}({exp2})
        ```

- and {condition1} {condition2} ... : 与条件块

    - condition1, condition2 ... : 若干条件块

    - ```
        ({condition1} and {condition2} and ... and {condition_last})
        ```

- or {condition1} {condition2} ... : 或条件块

    - condition1, condition2 ... : 若干条件块

    - ```
        ({condition1} or {condition2} or ... or {condition_last})
        ```

- not {condition} : 非条件块

    - condition: 条件块

    - ```
        (not {condition})
        ```

        