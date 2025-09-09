// Blockly块定义模板
const blockTemplates = {
  //赋值操作
  setto :{
    "message0":"赋值：\n将%2赋值给%1",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "exp",
        "check":"String",
      }],
    "previousStatement": null,
    "nextStatement": null,
    "colour":330,
    "tooltip": "将表达式的结果赋值给变量\nVAR: 目标变量名\nexp: 要赋值的表达式，可以是计算式、函数调用或其他值"
  },
  //工作簿操作
  openWorkbook: {
    type: "openWorkbook",
    message0: "打开工作簿：\n打开路径(%1)下的工作簿，将其命名为%2",
    args0: [
      {
        type: "field_input",
        name: "FILE_PATH",
        check: "String",
      },
      {
        type: "field_input",
        name: "WORKBOOK_VAR",
        check: "String",
      }
    ],
    nextStatement: null,
    previousStatement: null,
    colour: 200,
    tooltip: 'Open Workbook {path} As {var} : 打开指定路径的Excel工作簿并赋值给变量\n路径: Excel文档路径\n变量名: 工作簿对象的变量名称'
  },
  addWorkbook :{
    "message0":"新建工作簿：\n新建一个工作簿%1",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      }],
    "previousStatement": null,
    "nextStatement": null,
    "colour":200,
    "tooltip":'Add Workbook as {var} : 保存 Excel 文档 \
    \nworkbook: Excel 文档变量名 \
    \npath: 目标保存路径，为空表示在文档原位置覆盖保存'
  },
  saveWorkbook: {
    type: "saveWorkbook",
    message0: "保存工作簿：\n将工作簿%1保存到路径(%2)",
    args0: [
      {
        type: "field_input",
        name: "VAR",
        check: "String",
      },
      {
        type: "field_input",
        name: "FILE",
        check: "String",
      }
    ],
    nextStatement: null,
    previousStatement: null,
    colour: 200,
    tooltip: '保存工作簿到指定路径,提供默认路径，VAR为工作簿变量名，FILE为保存路径'
  },
  //行列操作
//获取行列
  fetchRowNoheader :{
    "message0":"获取行：\n获取工作簿%1中第%2行的第%3-%4列，存储到变量%5中",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "name":"row",
        "check":"number",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"columnF",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"columnT",
      },
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":"{workbook} Fetch Row {row} {column_from} {column_to}{header_row} As {var} : 获取一行\
    \nworkbook: Excel 文档变量名\
    \nrow: 行号，为空则采用当前活跃行\
    \ncolumn_from: 起点列号，为空则采用第一列\
    \ncolumn_to: 终点列号，为空则读取到空值为止\
    \nheader_row: header 所在行号，为空表示不需要 header\
    \nvar: 表示获取结果的变量名"
  },
  fetchRow :{
    "message0":"获取行（有表头）：\n获取工作簿%1中第%2行的第%3-%4列，存储到变量%6中，表头行号为%5",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "name":"row",
        "check":"number",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"columnF",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"columnT",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"header_row",
      },
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":"{workbook} Fetch Row {row} {column_from} {column_to}{header_row} As {var} : 获取一行\
    \nworkbook: Excel 文档变量名\
    \nrow: 行号，为空则采用当前活跃行\
    \ncolumn_from: 起点列号，为空则采用第一列\
    \ncolumn_to: 终点列号，为空则读取到空值为止\
    \nheader_row: header 所在行号，为空表示不需要 header\
    \nvar: 表示获取结果的变量名"
  },
  fetchCol :{
    "message0":"获取列：\n获取工作簿%1中第%2列的第%3-%4行，存储到变量%5中",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "name":"col",
        "check":"number",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"rowF",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"rowT",
      },
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":"{workbook} Fetch Column {column} {row_from} {row_to} As {var} : 获取一列\
    \nworkbook: Excel 文档变量名\
    \ncolumn: 列号，为空则采用当前活跃列\
    \nrow_from: 起点行号，为空则采用第一行\
    \nrow_to: 终点行号，为空则读取到空值为止\
    \nvar: 表示获取结果的变量名"
  },
//写入行列
  WriteRowNoheader: {
    type: "WriteRowNoheader",
    message0: "写入行：\n写入工作簿%1中第%2行的第%4列-%5列，行值为%3",
    args0: [
      {
        type: "field_input",
        name: "Workbook",
        check: "String",
      },
      {
        type: "field_input",
        name: "row",
        check: "Number",
      },
      {
        type: "field_input",
        name: "row_content",
        check: "String",
      },
      {
        type: "field_input",
        name: "column_from",
        check: "Number",
      },
      {
        type: "field_input",
        name: "column_to",
        check: "Number",
      },
    ],
    nextStatement: null,
    previousStatement: null,
    colour: 160,
    tooltip: "写入行：向工作簿中指定行写入内容，不支持表头行号",
  },
  WriteRow: {
    type: "WriteRow",
    message0: "写入行（有表头）：\n写入工作簿%1中第%2行的第%5列-%6列，行值为%3，表头行号为%4",
    args0: [
      {
        type: "field_input",
        name: "Workbook",
        check: "String",
      },
      {
        type: "field_input",
        name: "row",
        check: "Number",
      },
      {
        type: "field_input",
        name: "row_content",
        check: "String",
      },
      {
        type: "field_input",
        name: "header_row",
        check: "Number",
      },
      {
        type: "field_input",
        name: "column_from",
        check: "Number",
      },
      {
        type: "field_input",
        name: "column_to",
        check: "Number",
      },
    ],
    nextStatement: null,
    previousStatement: null,
    colour: 160,
    tooltip: "写入行（有表头）：向工作簿中指定行写入内容，支持表头行号",
  },
  WriteCol:{
    "message0":"写入列：\n写入工作簿%1的第%4行-%5行的第%2列，列值为%3",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"column",
      },
      {
        "type": "field_input",
        "check":"string",
        "name":"col_content",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"row_from",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"row_to",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
  },
//插入行列
  InsertColBefore: {
    type: "InsertColBefore",
    message0: "往前插入列：\n在工作簿%1中第%2列前插入%3列新列",
    args0: [
      {
        type: "field_input",
        name: "Workbook",
        check: "String",
      },
      {
        type: "field_input",
        name: "col",
        check: "Number",
      },
      {
        type: "field_input",
        name: "num_columns",
        check: "Number",
      },
    ],
    nextStatement: null,
    previousStatement: null,
    colour: 200,
    tooltip: "往前插入列：在指定列前插入新列",
  },
  InsertRowBefore: {
    type: "InsertRowBefore",
    message0: "往前插入行：\n在工作簿%1中第%2行前插入%3行新行",
    args0: [
      {
        type: "field_input",
        name: "Workbook",
        check: "String",
      },
      {
        type: "field_input",
        name: "row",
        check: "Number",
      },
      {
        type: "field_input",
        name: "num_rows",
        check: "Number",
      },
    ],
    nextStatement: null,
    previousStatement: null,
    colour: 200,
    tooltip: "往前插入行：在指定行前插入新行",
  },
//其他表格操作
  fetchCell: {
    type: "fetchCell",
    message0: "读取单元格：\n从工作簿%1读取单元格%2的值",
    args0: [
      {
        type: "field_input",
        name: "VAR",
        check: "String",
      },
      {
        type: "field_input",
        name: "CELL",
        check: "String",
      }
    ],
    nextStatement: null,
    previousStatement: null,
    colour: 200,
    tooltip: '读取工作簿中指定单元格的值，VAR为工作簿变量名，CELL为单元格引用'
  },
  fetchArea:{
    "message0":"获取区域：\n获取工作簿%1中第%2行-%3行，第%4列-%5列的全部内容\n以第%6行作为列名\n存储到变量%7",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"rowF",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"rowT",
      },
      {
        "type": "field_input",
        "name":"colF",
      },
      {
        "type": "field_input",
        "name":"colT",
      },
      {
        "type": "field_input",
        "name": "header",
      },
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
        "text":"SecondExcel",
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
  },
  SetCellValue: {
    type: "SetCellValue",
    message0: "设置单元格：\n在工作簿%1的单元格%2中写入值%3",
    args0: [
      {
        type: "field_input",
        name: "VAR",
        check: "String",
      },
      {
        type: "field_input",
        name: "CELL",
        check: "String",
      },
      {
        type: "field_input",
        name: "VALUE",
        check: "String",
      }
    ],
    nextStatement: null,
    previousStatement: null,
    colour: 200,
    tooltip: '向工作簿中指定单元格写入值，VAR为工作簿变量名，CELL为单元格引用，VALUE为要写入的值'
  },
  SetAreaText:{
    "message0":"设置区域数据格式为文本：\n把工作簿%1中第%2行-第%3行,第%4列-第%5列内的数据改为纯文本类型",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "name":"row_from",
        "check":"number",

      },
      {
        "type": "field_input",
        "name":"row_to",
        "check":"number",

      },
      {
        "type": "field_input",
        "name":"column_from",
        "check":"number",

      },
      {
        "type": "field_input",
        "name":"column_to",
        "check":"number",

      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
  },
//工作表操作
  CreateSheet:{
    "message0":"新建工作表:\n在工作簿%1里新建名为%2的工作表",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "name",
        "check":"String",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":240,
    "tooltip":'{workbook} Create Sheet {name} : 新建表\
    \nworkbook: Excel 文档变量名\
    \nname: 新建的表名'
  },
  SetActiveSheet:{
    "message0":"设置活跃工作表：\n在工作簿%1里设置名为%2的工作表为活跃工作表",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "name",
        "check":"String",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":240,
    "tooltip":'{workbook} Set Active Sheet {name} : 改变活跃中的表\
    \nworkbook: Excel 文档变量名\
    \nname: 将要设为活跃的表名'
  },
//表格集合操作
  MakeWorkbookDict:{
    "message0": "创建集合：\n创建表格集合%1",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      }
    ],
    "nextStatement": null,
    "previousStatement":null,
    "colour":300,
    "tooltip":"Make workbook dictionary {var}: 生成一个workbook集合\
    \nvar: 生成的workbook集合变量名"
  },
  SetDictHeaders:{
    "message0": "设置集合表头：\n将表格集合%1的第%2行设置为表头，表头内容为%3\n",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"header_row",
      },
      {
        "type": "field_input",
        "check":"String",
        "name":"headers",
      },
    ],
    "nextStatement": null,
    "previousStatement":null,
    "colour":300,
    "tooltip":'{workbook_dict} set headers {headers} {header_row}: 设置workbook集合的表头，设置后将会按照对应表头插入内容 \
    \nworkbook_dict: workbook集合\
    \n headers: 设置的表头\
    \nheader_row: headers所在行号'
  },
  AddRowDict:{
    "message0": "集合添加新行：\n在表格集合%1中找到工作簿%2，新增一行,内容为:%3\n",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"name",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"row_content",
      },
    ],
    "nextStatement": null,
    "previousStatement":null,
    "colour":300,
    "tooltip":'{workbook_dict} add row {name} {row_content}: 向一个workbook新增一行\
    \nworkbook_dict: workbook集合\
    \nname: 需要新增一行的workbook名\
    \nrow_content: 新增内容'
  },
  SetDictColText :{
    "message0":"设置集合列数据格式为文本：\n把表格集合%1中第%2列内的数据改为纯文本类型",
    "args0": [
      {
        "type": "field_input",
        "name": "Dict",
        "check":"String",
      },
      {
        "type": "field_input",
        "name":"column",
        "check":"number",

      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":300,
  },
  GenerateFile:{
    "message0": "导出工作簿：\n将表格集合%1中的工作簿导出到路径%2",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "Path",
        "check":"String",
      },

    ],
    "nextStatement": null,
    "previousStatement":null,
    "colour":300,
    "tooltip":'{workbook_dict} generate workbook files {path}: 生成excel文件\
    \nworkbook_dict: workbook集合\
    \npath: 生成目录，默认为 \'./\''
  },
  getAllWorkbook :{
    "type":"getAllWorkbook",
    "message0": "获取工作簿名：\n获取表格集合%1中所有工作簿的名信息，整合到变量%2中",
    "args0": [
      {
        "type": "field_input",
        "name": "Dict",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      }
    ],
    "nextStatement": null,
    "previousStatement":null,
    "colour":300,
    "tooltip": "获取指定表格集合中所有工作簿的名称信息\nDict: 表格集合的变量名\nVAR: 存储工作簿名称列表的变量名"
  },
//辅助功能
  SettoNum :{
    "message0":"字符串列号转换为数字列号：\n将字符串%2转换成数字赋值给%1",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "exp",
        "check":"String",
      }],
    "previousStatement": null,
    "nextStatement": null,
    "colour":350,
  },
  SettoString:{
    "message0":"数字列号转换为字符串列号：\n将数字%2转换成字符串赋值给%1",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "exp",
        "check":"number",
      }],
    "previousStatement": null,
    "nextStatement": null,
    "colour":350,
  },
  GetStringLength :{
    "message0":"获取字符串长度：\n获取字符串%1的长度，存储到变量%2",
    "args0": [
      {
        "type": "field_input",
        "name": "str",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      }],
    "previousStatement": null,
    "nextStatement": null,
    "colour":220,
  },
  SubString :{
    "message0":"截取字符串片段：\n截取字符串%1的第%2个字符到第%3个字符，存储到变量%4",
    "args0": [
      {
        "type": "field_input",
        "name": "str",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "head",
        "check":"number",
      },
      {
        "type": "field_input",
        "name": "tail",
        "check":"number",
      },
      {
        "type": "field_input",
        "name": "VAR",
        "check":"number",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":220,
  },
  TeleVerify :{
    "message0":"对字符串进行手机号码规则校验：\n判断字符串%1是否为手机号",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      }],
    "previousStatement": null,
    "nextStatement": null,
    "colour":220,
  },
  LeftJust :{
    "message0":"左对齐：\n获取字符串%1，将其补齐到%3位，存储到变量%2",
    "args0": [
      {
        "type": "field_input",
        "name": "str",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "Length",
        "check":"Number",
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":220,
  },
//条件语句
  IF :{
    "message0":"if %1 : %2",
    "args0": [
      {
        "type": "input_value",
        "name":"condition",
        
      },
      {
        "type": "input_statement",
        "name":"content",
        "check":null,
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":220,
    "returns":"loop",
    "tooltip":' if {condition} : 分支控制块 if'
  },
  Elif :{
    "message0":"Else if %1 : %2",
    "args0": [
      {
        "type": "input_statement",
        "name":"condition",
        "check":"string",
      },
      {
        "type": "input_statement",
        "name":"content",
        "check":"string",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":220,
    "tooltip":'else if {condition} : 分支控制块 else if'
  },
  Else :{
    "message0":"Else %1",
    "args0": [
      {
        "type": "input_statement",
        "name":"content",
        "check":"string",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":220,
    'tooltip':'分支控制块 else'
  },
//循环
  For :{
    "message0":"for %1 from %2 to %3 : %4",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
        "text":"i",
      },
      {
        "type": "field_input",
        "name":"start",
        "check":"number",

      },
      {
        "type": "field_input",
        "check":"number",
        "name":"end",
      },
      {
        "type": "input_statement",
        "name":"content",
        "check":"string",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":220,
    'tooltip':'for {var} from {start} to {end} : 循环控制块，带int型循环变量，前闭后闭，start <= end'
  },
  Foreach:{
    "message0":"for each %1  in %2 :%3",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String",
        "text":"i",
      },
      {
        "type": "field_input",
        "name":"iterable_var",
        "check":"string",

      },
      {
        "type": "input_statement",
        "name":"content",
        "check":"string",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":220,
    'tooltip':'for each {var} in {iterable_var} : 循环控制块，for each 循环，针对可迭代容器'
  },
  While :{
   "type":"conditionconnective",
    "message0":"while %1 :%2",
    "args0": [
      {
        "type": "input_statement",
        "name":"condition1",
      },
      {
        "type": "input_statement",
        "name":"content",
        "check":"string",
      },
    ],
    "output": null,
    "previousStatement":null,
    "nextStatement":null,
    "colour":220,
  },
//逻辑块
  Compare :{
    "message0":" %1 %2 %3 %4",
    "args0": [
      {
        "type": "field_dropdown",
        "name": "value_type",
        "options": [
          [ "整数", "int" ],
          [ "小数", "float" ],
          ["字符串","str"],
        ]
      },
      {
        "type": "field_input",
        "name":"exp1",
        "check":"string",
      },
      {
        "type": "field_dropdown",
        "name": "comparation",
        "options": [
          [ "小于", "<" ],
          [ "小于等于", "<=" ],
          ["等于","=="],
          ["大于等于",">="],
          ["大于",">"],
          ["不等于","!="]
        ]
      },
      {
        "type": "field_input",
        "name":"exp2",
        "check":"string",
      },
    ],
    "output":null,
    "previousStatement": null,
    "nextStatement":null,
    "colour":220,
    "tooltip":'{value_type} {exp1} {comparation} {exp2} : 比较条件块\
    \nvalue_type: 表达式数据类型，三种可选项为 int float str\
    \nexp1, exp2: 参与比较的两个表达式\
    \ncomparation: 比较运算符，五种可选项为 < <= == >= >'
  },
  isNone :{
    "message0":" %1 %2 空",
    "args0": [
      {
        "type": "field_input",
        "name":"VAR",
        "check":"string",
      },
      {
        "type": "field_dropdown",
        "name": "option",
        "options": [
          [ "为", " is" ],
          [ "不为", " is not" ],
        ]
      },
    ],
    "output":null,
    "previousStatement": null,
    "nextStatement":null,
    "colour":220,
  },
  and :{
    "type":"conditionconnective",
    "message0":"and %1",
    "args0": [
      {
        "type": "input_statement",
        "name":"condition1",
      },
    ],
    "output": null,
    "previousStatement":null,
    "nextStatement":null,
    "colour":220,
    'tooltip':'and {condition1} {condition2} ... : 与条件块'
  },
  or :{
    "message0":"or %1",
    "args0": [
      {
        "type": "input_statement",
        "name":"condition1",
      },
    ],
    "output": null,
    "previousStatement":null,
    "nextStatement":null,
    "colour":220,
    'tooltip':'or {condition1} {condition2} ... : 或条件块'
  },
  not :{
    "message0":"not %1",
    "args0": [
      {
        "type": "input_value",
        "name":"condition1",
      },
    ],
    "colour":220,
    "output": null,
    "previousStatement":null,
    "nextStatement":null,
    'tooltip':'not {condition} : 非条件块'
  }

};

const tools = [
  //赋值操作
  {
  type: "function",
  function: {
    name: "assign_value",
    description: "将表达式的结果赋值给变量",
    parameters: {
      type: "object",
      properties: {
        variable_name: {
          type: "string",
          description: "目标变量名，例如：'result'、'total'、'calculated_value'"
        },
        expression: {
          type: "string",
          description: "要赋值的表达式，可以是计算式、函数调用或其他变量的值，例如：'A1+B1'、'SUM(C2:C10)'、'previous_total * 1.1'"
        }
      },
      required: ["variable_name", "expression"]
    }
  }
},
 //工作簿操作
  {
    type: "function",
    function: {
      name: "open_workbook",
      description: "打开Excel工作簿文件，可以指定文件路径和变量名",
      parameters: {
        type: "object",
        properties: {
          file_path: {
            type: "string",
            description: "Excel文档的路径，例如：'C:/文档/销售数据.xlsx'、'./data/财务报表.xlsx'"
          },
          variable_name: {
            type: "string",
            description: "表示文档的变量名，用于后续操作引用，例如：'sales_data'、'financial_report'"
          }
        },
        required: ["file_path", "variable_name"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "add_workbook", // 新增工具定义
      description: "创建一个新的Excel工作簿",
      parameters: {
        type: "object",
        properties: {
          variable_name: {
            type: "string",
            description: "表示新工作簿的变量名，用于后续操作引用，例如：'new_workbook'、'data_report'"
          }
        },
        required: ["variable_name"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "save_workbook",
      description: "保存工作簿到指定路径",
      parameters: {
        type: "object",
        properties: {
          file_path: {
            type: "string",
            description: "保存的文件路径，例如：'C:/文档/报表.xlsx'、'./output/结果.csv'"
          },
          workbook_variable: {
            type: "string",
            description: "要保存的工作簿变量名，例如：'sales_data'、'current_workbook'"
          }
        },
        required: ["file_path", "workbook_variable"]
      }
    }
  },
//行列操作
//获取行列
  {
  type: "function",
  function: {
    name: "fetch_row_no_header",
    description: "获取工作簿中指定行的内容（无表头）",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'、'current_workbook'"
        },
        row: {
          type: "number",
          description: "行号，例如：1、2、3"
        },
        column_from: {
          type: "number",
          description: "起始列号，例如：1、2"
        },
        column_to: {
          type: "number",
          description: "结束列号，例如：5、10"
        },
        variable_name: {
          type: "string",
          description: "存储结果的变量名，例如：'row_data'、'extracted_row'"
        }
      },
      required: ["workbook_variable", "row", "column_from", "column_to", "variable_name"]
    }
  }
  },
  {
  type: "function",
  function: {
    name: "fetch_row_with_header",
    description: "从工作簿中获取指定行的内容（包含表头信息）",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'、'current_workbook'"
        },
        row: {
          type: "number",
          description: "行号，例如：1、2、3"
        },
        column_from: {
          type: "number",
          description: "起始列号，例如：1、2"
        },
        column_to: {
          type: "number",
          description: "结束列号，例如：5、10"
        },
        header_row: {
          type: "number",
          description: "表头行号，例如：1"
        },
        variable_name: {
          type: "string",
          description: "存储结果的变量名，例如：'row_data'、'extracted_row'"
        }
      },
      required: ["workbook_variable", "row", "column_from", "column_to", "header_row", "variable_name"]
    }
  }
  },
  {
  type: "function",
  function: {
    name: "fetch_column",
    description: "获取工作簿中指定列的内容",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'、'current_workbook'"
        },
        column: {
          type: "number",
          description: "列号，例如：1、2、3（A=1, B=2, C=3...）"
        },
        row_from: {
          type: "number",
          description: "起始行号，例如：1、2"
        },
        row_to: {
          type: "number",
          description: "结束行号，例如：10、20"
        },
        variable_name: {
          type: "string",
          description: "存储结果的变量名，例如：'column_data'、'extracted_column'"
        }
      },
      required: ["workbook_variable", "column", "row_from", "row_to", "variable_name"]
    }
  }
  },
//写入行列
  {
  type: "function",
  function: {
    name: "write_row_no_header",
    description: "向工作簿中写入指定行的内容（不包含表头）",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'"
        },
        row: {
          type: "number",
          description: "行号，例如：1、2、3"
        },
        row_content: {
          type: "string",
          description: "写入的行内容，例如：'100,200,300'"
        },
        column_from: {
          type: "number",
          description: "起始列号，例如：1、2"
        },
        column_to: {
          type: "number",
          description: "结束列号，例如：5、10"
        }
      },
      required: ["workbook_variable", "row", "row_content", "column_from", "column_to"]
    }
  }
  },
  {
    type: "function",
    function: {
      name: "write_row",
      description: "向工作簿中写入指定行的内容",
      parameters: {
        type: "object",
        properties: {
          workbook_variable: {
            type: "string",
            description: "工作簿变量名，例如：'sales_data'"
          },
          row: {
            type: "number",
            description: "行号，例如：1、2、3"
          },
          row_content: {
            type: "string",
            description: "写入的行内容，例如：'100,200,300'"
          },
          header_row: {
            type: "number",
            description: "表头行号，例如：1"
          },
          column_from: {
            type: "number",
            description: "起始列号，例如：1、2"
          },
          column_to: {
            type: "number",
            description: "结束列号，例如：5、10"
          }
        },
        required: ["workbook_variable", "row", "row_content", "column_from", "column_to"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "write_column",
      description: "向工作簿中写入指定列的内容",
      parameters: {
        type: "object",
        properties: {
          workbook_variable: {
            type: "string",
            description: "工作簿变量名，例如：'sales_data'、'current_workbook'"
          },
          column: {
            type: "number",
            description: "列号，例如：1、2、3（A=1, B=2, C=3...）"
          },
          column_content: {
            type: "string",
            description: "要写入的列内容，可以是逗号分隔的值，例如：'100,200,300,400'"
          },
          row_from: {
            type: "number",
            description: "起始行号，例如：1、2"
          },
          row_to: {
            type: "number",
            description: "结束行号，例如：10、20"
          }
        },
        required: ["workbook_variable", "column", "column_content", "row_from", "row_to"]
      }
    }
  },
//插入行列
  {
    type: "function",
    function: {
      name: "insert_column",
      description: "在工作簿中插入新列",
      parameters: {
        type: "object",
        properties: {
          workbook_variable: {
            type: "string",
            description: "工作簿变量名，例如：'sales_data'"
          },
          column: {
            type: "number",
            description: "列号，例如：1、2、3"
          },
          num_columns: {
            type: "number",
            description: "插入的列数，例如：1、2"
          }
        },
        required: ["workbook_variable", "column", "num_columns"]
      }
    }
  },
{
  type: "function",
  function: {
    name: "insert_row",
    description: "在工作簿中插入新行",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'"
        },
        row: {
          type: "number",
          description: "行号，例如：1、2、3"
        },
        num_rows: {
          type: "number",
          description: "插入的行数，例如：1、2"
        }
      },
      required: ["workbook_variable", "row", "num_rows"]
    }
  }
},
//其他表格操作
{
  type: "function",
  function: {
    name: "fetch_cell",
    description: "从工作簿中读取指定单元格的值",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'"
        },
        cell_reference: {
          type: "string",
          description: "单元格引用，例如：'A1'、'B5'"
        }
      },
      required: ["workbook_variable", "cell_reference"]
    }
  }
},
{
  type: "function",
  function: {
    name: "fetch_area",
    description: "从工作簿中获取指定区域的内容",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'"
        },
        row_from: {
          type: "number",
          description: "起始行号，例如：1、2"
        },
        row_to: {
          type: "number",
          description: "结束行号，例如：10、20"
        },
        column_from: {
          type: "number",
          description: "起始列号，例如：1、2"
        },
        column_to: {
          type: "number",
          description: "结束列号，例如：5、10"
        },
        header_row: {
          type: "number",
          description: "表头行号，例如：1"
        },
        variable_name: {
          type: "string",
          description: "存储结果的变量名，例如：'area_data'"
        }
      },
      required: ["workbook_variable", "row_from", "row_to", "column_from", "column_to", "header_row", "variable_name"]
    }
  }
},
{
  type: "function",
  function: {
    name: "set_cell_value",
    description: "向工作簿中指定单元格写入值",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'"
        },
        cell_reference: {
          type: "string",
          description: "单元格引用，例如：'A1'、'B5'"
        },
        value: {
          type: "string",
          description: "要写入的值，例如：'100'、'Hello World'"
        }
      },
      required: ["workbook_variable", "cell_reference", "value"]
    }
  }
},
{
  type: "function",
  function: {
    name: "set_area_text",
    description: "将工作簿中指定区域的数据格式设置为文本",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'"
        },
        row_from: {
          type: "number",
          description: "起始行号，例如：1、2"
        },
        row_to: {
          type: "number",
          description: "结束行号，例如：10、20"
        },
        column_from: {
          type: "number",
          description: "起始列号，例如：1、2"
        },
        column_to: {
          type: "number",
          description: "结束列号，例如：5、10"
        }
      },
      required: ["workbook_variable", "row_from", "row_to", "column_from", "column_to"]
    }
  }
},

//工作表操作
{
  type: "function",
  function: {
    name: "create_sheet",
    description: "在工作簿中创建一个新的工作表",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'"
        },
        sheet_name: {
          type: "string",
          description: "新工作表的名称，例如：'Sheet1'、'Summary'"
        }
      },
      required: ["workbook_variable", "sheet_name"]
    }
  }
},
{
  type: "function",
  function: {
    name: "set_active_sheet",
    description: "设置工作簿中的活跃工作表",
    parameters: {
      type: "object",
      properties: {
        workbook_variable: {
          type: "string",
          description: "工作簿变量名，例如：'sales_data'"
        },
        sheet_name: {
          type: "string",
          description: "要设置为活跃的工作表名称，例如：'Sheet1'、'Summary'"
        }
      },
      required: ["workbook_variable", "sheet_name"]
    }
  }
},
//表格集合操作
{
  type: "function",
  function: {
    name: "make_workbook_dict",
    description: "创建一个表格集合",
    parameters: {
      type: "object",
      properties: {
        variable_name: {
          type: "string",
          description: "表格集合的变量名，例如：'workbook_dict'"
        }
      },
      required: ["variable_name"]
    }
  }
},
{
  type: "function",
  function: {
    name: "set_dict_headers",
    description: "设置表格集合的表头",
    parameters: {
      type: "object",
      properties: {
        workbook_dict: {
          type: "string",
          description: "表格集合的变量名，例如：'workbook_dict'"
        },
        header_row: {
          type: "number",
          description: "表头所在行号，例如：1"
        },
        headers: {
          type: "string",
          description: "表头内容，例如：'Name, Age, Salary'"
        }
      },
      required: ["workbook_dict", "header_row", "headers"]
    }
  }
},
{
  type: "function",
  function: {
    name: "add_row_dict",
    description: "向表格集合中的指定工作簿添加新行",
    parameters: {
      type: "object",
      properties: {
        workbook_dict: {
          type: "string",
          description: "表格集合的变量名，例如：'workbook_dict'"
        },
        workbook_name: {
          type: "string",
          description: "目标工作簿的名称，例如：'sales_data'"
        },
        row_content: {
          type: "string",
          description: "新增行的内容，例如：'John, 30, 5000'"
        }
      },
      required: ["workbook_dict", "workbook_name", "row_content"]
    }
  }
},
{
  type: "function",
  function: {
    name: "set_dict_col_text",
    description: "将表格集合中的指定列数据格式设置为文本",
    parameters: {
      type: "object",
      properties: {
        workbook_dict: {
          type: "string",
          description: "表格集合的变量名，例如：'workbook_dict'"
        },
        column: {
          type: "number",
          description: "列号，例如：1、2、3"
        }
      },
      required: ["workbook_dict", "column"]
    }
  }
},
{
  type: "function",
  function: {
    name: "generate_file",
    description: "导出表格集合中的工作簿到指定路径",
    parameters: {
      type: "object",
      properties: {
        workbook_dict: {
          type: "string",
          description: "表格集合的变量名，例如：'workbook_dict'"
        },
        path: {
          type: "string",
          description: "导出路径，例如：'C:/output/'"
        }
      },
      required: ["workbook_dict", "path"]
    }
  }
},
  {
  type: "function",
  function: {
    name: "get_all_workbooks",
    description: "获取表格集合中所有工作簿的名称信息",
    parameters: {
      type: "object",
      properties: {
        workbook_collection: {
          type: "string",
          description: "表格集合的变量名，例如：'excel_app'、'workbooks_dict'"
        },
        variable_name: {
          type: "string",
          description: "存储结果的变量名，例如：'workbook_names'、'sheet_list'"
        }
      },
      required: ["workbook_collection", "variable_name"]
    }
  }
  },


];



// 函数名到Blockly块类型的映射
const functionToBlockMap = {
  "assign_value": "setto",

  "open_workbook": "openWorkbook",
  "add_workbook": "addWorkbook",
  "save_workbook": "saveWorkbook",
  
  "fetch_row_no_header": "fetchRowNoheader",
  "fetch_row_with_header": "fetchRow",
  "fetch_column": "fetchCol",

  "write_row_no_header": "WriteRowNoheader",
  "write_row": "WriteRow",
  "write_column": "WriteCol",

  "insert_column": "InsertColBefore",
  "insert_row": "InsertRowBefore",

  "fetch_cell": "fetchCell",
  "fetch_area": "fetchArea",
  "set_cell_value": "SetCellValue",
  "set_area_text": "SetAreaText",
  
  "create_sheet": "CreateSheet",
  "set_active_sheet": "SetActiveSheet",

  "make_workbook_dict": "MakeWorkbookDict",
  "set_dict_headers": "SetDictHeaders",
  "add_row_dict": "AddRowDict",
  "set_dict_col_text": "SetDictColText",
  "generate_file": "GenerateFile",
  "get_all_workbooks": "getAllWorkbook",


};

// 参数映射 - 将AI返回的参数名映射到Blockly块中的字段名
const paramMapping = {
  //赋值操作
  "assign_value": {
  "variable_name": "VAR",
  "expression": "exp"
},
//工作簿操作
  "open_workbook": {
    "file_path": "FILE_PATH",
    "variable_name": "WORKBOOK_VAR"
  },
  "add_workbook": { 
    "variable_name": "VAR"
  },
  "save_workbook": {
    "file_path": "FILE",
    "workbook_variable": "VAR"
  },
//行列操作
//获取行列
  "fetch_row_no_header": {
  "workbook_variable": "Workbook",
  "row": "row",
  "column_from": "columnF",
  "column_to": "columnT",
  "variable_name": "VAR"
},
  "fetch_row_with_header": {
  "workbook_variable": "Workbook",
  "row": "row",
  "column_from": "columnF",
  "column_to": "columnT",
  "header_row": "header_row",
  "variable_name": "VAR"
},
  "fetch_column": {
  "workbook_variable": "Workbook",
  "column": "col",
  "row_from": "rowF",
  "row_to": "rowT",
  "variable_name": "VAR"
},
//写入行列
  "write_row_no_header": {
  "workbook_variable": "Workbook",
  "row": "row",
  "row_content": "row_content",
  "column_from": "column_from",
  "column_to": "column_to"
},
  "write_row": {
    "workbook_variable": "Workbook",
    "row": "row",
    "row_content": "row_content",
    "header_row": "header_row",
    "column_from": "column_from",
    "column_to": "column_to"
  },
  "write_column": {
  "workbook_variable": "Workbook",
  "column": "column",
  "column_content": "col_content",
  "row_from": "row_from",
  "row_to": "row_to"
},
//插入行列
  "insert_column": {
    "workbook_variable": "Workbook",
    "column": "col",
    "num_columns": "num_columns"
  },
  "insert_row": {
    "workbook_variable": "Workbook",
    "row": "row",
    "num_rows": "num_rows"
  },
//其他表格操作
  "fetch_cell": {
  "workbook_variable": "VAR",
  "cell_reference": "CELL"
},
"fetch_area": {
  "workbook_variable": "Workbook",
  "row_from": "rowF",
  "row_to": "rowT",
  "column_from": "colF",
  "column_to": "colT",
  "header_row": "header",
  "variable_name": "VAR"
},
"set_cell_value": {
  "workbook_variable": "VAR",
  "cell_reference": "CELL",
  "value": "VALUE"
},
"set_area_text": {
  "workbook_variable": "Workbook",
  "row_from": "row_from",
  "row_to": "row_to",
  "column_from": "column_from",
  "column_to": "column_to"
},
//工作表操作
"create_sheet": {
  "workbook_variable": "Workbook",
  "sheet_name": "name"
},
"set_active_sheet": {
  "workbook_variable": "Workbook",
  "sheet_name": "name"
},
//表格集合操作
"make_workbook_dict": {
  "variable_name": "VAR"
},
"set_dict_headers": {
  "workbook_dict": "VAR",
  "header_row": "header_row",
  "headers": "headers"
},
"add_row_dict": {
  "workbook_dict": "VAR",
  "workbook_name": "name",
  "row_content": "row_content"
},
"set_dict_col_text": {
  "workbook_dict": "Dict",
  "column": "column"
},
"generate_file": {
  "workbook_dict": "VAR",
  "path": "Path"
},
  "get_all_workbooks": {
  "workbook_collection": "Dict",
  "variable_name": "VAR"
},


};

async function getAIResponse(userInput) {
  const apiKey = 'sk-8cd20324714d415c8dffc8f9b2fedf25';
  const baseURL = 'https://api.deepseek.com';

  try {
    const response = await fetch(`${baseURL}/v1/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: 'deepseek-chat',
        messages: [
          {
            role: 'system',
            content: '你是一个Excel操作助手，可以将用户请求转换为复杂的Excel操作函数。请根据用户请求选择合适的函数并正确填充参数。'
          },
          {
            role: 'user',
            content: userInput
          }
        ],
        tools: tools, // 添加工具定义
        tool_choice: 'auto', // 让AI自动选择工具
        temperature: 0.1, // 降低温度以获得更确定的响应
        max_tokens: 2000
      })
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(`HTTP ${response.status}: ${errorData.error?.message || response.statusText}`);
    }

    const data = await response.json();
    
    if (!data.choices || !Array.isArray(data.choices) || data.choices.length === 0) {
      throw new Error('API响应格式异常: choices字段缺失或为空');
    }
    
    return data.choices[0].message;
  } catch (error) {
    console.error('调用 DeepSeek API 时出错:', error);
    throw error;
  }
}

async function handleUserCommand(userInput) {
  try {
    const aiResponse = await getAIResponse(userInput);
    
    // 检查是否有工具调用
    if (aiResponse.tool_calls && aiResponse.tool_calls.length > 0) {
      console.log("检测到函数调用，生成Blockly块...");
      
      for (const toolCall of aiResponse.tool_calls) {
        const functionName = toolCall.function.name;
        const functionArgs = JSON.parse(toolCall.function.arguments);
        
        console.log(`函数: ${functionName}`);
        console.log(`参数: ${JSON.stringify(functionArgs)}`);
        
        // 转换为Blockly块配置
        const blockConfig = convertFunctionToBlockConfig(functionName, functionArgs);
        if (blockConfig) {
          generateBlock(blockConfig);
        }
      }
      return true;
    } else if (aiResponse.content) {
      // 处理普通文本响应
      console.log("AI 回复:", aiResponse.content);
      
      // 这里可以添加将文本响应显示给用户的逻辑
      displayResponseToUser(aiResponse.content);
      return false;
    } else {
      console.log("AI返回了空响应");
      return false;
    }
  } catch (error) {
    console.error("处理用户命令时出错:", error);
    displayResponseToUser("处理命令时出错: " + error.message);
    return false;
  }
}

function convertFunctionToBlockConfig(functionName, functionArgs) {
  const blockType = functionToBlockMap[functionName];
  if (!blockType) {
    console.warn(`未知函数: ${functionName}`);
    return null;
  }
  
  const mapping = paramMapping[functionName];
  const params = {};
  
  for (const [aiParam, blockParam] of Object.entries(mapping)) {
    if (functionArgs[aiParam] !== undefined) {
      params[blockParam] = functionArgs[aiParam];
    }
  }
  // 嵌套处理
  if (functionArgs.true_content || functionArgs.false_content) {
    params.content_true = functionArgs.true_content.map((nestedFunction) =>
      convertFunctionToBlockConfig(nestedFunction.name, nestedFunction.arguments)
    );
    params.content_false = functionArgs.false_content.map((nestedFunction) =>
      convertFunctionToBlockConfig(nestedFunction.name, nestedFunction.arguments)
    );
  }
  return {
    type: blockType,
    params: params
  };
}
function displayResponseToUser(message) {
  // 将消息显示在AI回复文本框中
  const responseTextarea = document.getElementById('ai-response');
  if (responseTextarea) {
    responseTextarea.value = message;
    
    // 自动滚动到底部
    responseTextarea.scrollTop = responseTextarea.scrollHeight;
  } else {
    console.error("找不到AI回复文本框元素");
    // 备用方案：在控制台显示
    console.log("显示给用户:", message);
  }
}

function createResponseElement() {
  const div = document.createElement('div');
  div.id = 'ai-response';
  div.style.position = 'fixed';
  div.style.bottom = '10px';
  div.style.right = '10px';
  div.style.background = '#f0f0f0';
  div.style.padding = '10px';
  div.style.borderRadius = '5px';
  div.style.boxShadow = '0 2px 10px rgba(0,0,0,0.1)';
  div.style.maxWidth = '300px';
  div.style.zIndex = '1000';
  document.body.appendChild(div);
  return div;
}
function generateBlock(config) {
  
  // 获取块模板
  const blockTemplate = blockTemplates[config.type];
  if (!blockTemplate) {
    console.error(`找不到块模板: ${config.type}`);
    return;
  }
  
  // 这里是实际的Blockly块生成逻辑
  // 1. 创建 Blockly 块实例
  const block = workspace.newBlock(config.type);
  
  // 2. 设置参数
  for (const [key, value] of Object.entries(config.params)) {
    if (block.getField(key)) {
      block.setFieldValue(value, key);
    } else if (Array.isArray(value)) {
      // 如果参数是嵌套块列表，递归生成嵌套块
      value.forEach((nestedConfig) => generateBlock(nestedConfig));
    } 
    else {
      console.warn(`字段 "${key}" 在块 "${config.type}" 中未找到`);
    }
  }
  
  // 3. 添加到工作区
  block.initSvg();
  block.render();
  
  // 4. 定位到可视区域
  workspace.centerOnBlock(block.id);
  
  // return blockTemplate;
  const blockInfo = {
    type: config.type,
    fields: config.params
  };
  
  console.log("生成的块信息:", blockInfo);
  return blockInfo;
}

// const testCommand = `
//   打开路径'C:/文档/销售数据.xlsx'下的工作簿，将其命名为sales_data；
//   从工作簿'sales_data'中读取第1行的内容，从第1列到第5列，存储为变量'row_data'，表头为第1行；
//   向工作簿'sales_data'的第2行写入内容'100,200,300'，从第1列到第3列；
//   保存工作簿'sales_data'到路径'C:/文档/结果.xlsx'
// `;

// // // 测试所有示例
// console.log(`\n测试命令: "${testCommand}"`);
// handleUserCommand(testCommand)
//   .then((result) => {
//     console.log("测试命令执行成功:", result);
//   })
//   .catch((error) => {
//     console.error("测试命令执行失败:", error);
//   });
