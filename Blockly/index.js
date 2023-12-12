'use strict';
let workspace = null;
var depth=1;
function start() {
  registerFirstContextMenuOptions();

 
  registerOpenWorkbook();
  registerSaveWorkbook();

  registerMoveActiveCell();
  registerSetActiveCell();
  registerFetchCell();
  registerFetchRow();
  registerFetchCol();
  registerFetchArea();
  registerInsertRow();
  registerInsertCol();
  registerSetCellValue();

  registerCreateSheet();
  registerSetActiveSheet();
  registerMergeSheet();
  registerCompareBlock();
  registerIfBlock();
  registerElifBlock();
  registerElseBlock();

  registerSettoBlock();

  registerForBlock();
  registerForeachBlock();

  registerAndBlock();
  registerOrBlock();
  registerNotBlock();

  workspace = Blockly.inject('blocklyDiv',
    {
        toolbox:document.getElementById('toolbox-categories'),
        grid:
         {spacing: 20,
          length: 5,
          colour: "#000",
          snap: true},
      trashcan: true
    });
  // workspace.addChangeListener(event => {
  //     const code = python.pythonGenerator.workspaceToCode(workspace);
  //     document.getElementById('generatedCodeContainer').value = code;
  //   });
  registerOutputOption();
  registerHelpOption();
  registerDisplayOption();
  Blockly.ContextMenuRegistry.registry.unregister('workspaceDelete');
}
function registerFirstContextMenuOptions() {
  // This context menu item shows how to use a precondition function to set the visibility of the item.
  const workspaceItem = {
    displayText: 'Hello World',
    // Precondition: Enable for the first 30 seconds of every minute; disable for the next 30 seconds.
    preconditionFn: function(scope) {
      const now = new Date(Date.now());
      if (now.getSeconds() < 30) {
        return 'enabled';
      }
      return 'disabled';
    },
    callback: function(scope) {
    },
    scopeType: Blockly.ContextMenuRegistry.ScopeType.WORKSPACE,
    id: 'hello_world',
    weight: 100,
  };
  // Register.
  Blockly.ContextMenuRegistry.registry.register(workspaceItem);
  
  // Duplicate the workspace item (using the spread operator).
  const blockItem = {...workspaceItem}
  // Use block scope and update the id to a nonconflicting value.
  blockItem.scopeType = Blockly.ContextMenuRegistry.ScopeType.BLOCK;
  blockItem.id = 'hello_world_block';
  Blockly.ContextMenuRegistry.registry.register(blockItem);
}
function registerOpenWorkbook()
{ 
  var openWorkbook = {
    "type":"openWorkbook",
    "message0": "open Workbook %1 As %2",
    "args0": [
      {
        "type": "field_input",
        "name": "FILE",
        "check":"String"
      },
      {
        "type": "field_input",
        "name": "VAR",
        "variable": "item",
        "variableTypes": [""],
        "check":"String"
      }
    ],
    "nextStatement": null,
    "previousStatement":null,
    "colour":200,
    "tooltip":'Open Workbook {path} As {var} : 打开Excel文档 \
    \npath: Excel 文档路径，为空表示新建文档 \
    \nvar: 表示文档的变量名'
  };
  Blockly.Blocks['openWorkbook']=
    {
      init: function() {
        this.jsonInit(openWorkbook);
      } 
    };
    python.pythonGenerator.forBlock['openWorkbook'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
      var FILE =block.getFieldValue('FILE');
      var FILEPATH;
      if(FILE!='')
      FILEPATH = '\'' + FILE + '\'';
      else
        FILEPATH=FILE;
        var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=VAR+"=ExcelApplication()\n";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code +=VAR+".open_application(visible=True)\n";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
        code +=VAR+".open_workbook("+FILEPATH+")\n";
      return code;
    }       
          
}
function registerSaveWorkbook(){
  var saveWorkbook ={
    "type":"saveWorkbook",
    "message0":"%1 Save Workbook %2",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "name": "FILE",
        "check":"String"
      }],
    "previousStatement": null,
    "nextStatement": null,
    "colour":200,
    "tooltip":'{workbook} Save Workbook {path} : 保存 Excel 文档 \
    \nworkbook: Excel 文档变量名 \
    \npath: 目标保存路径，为空表示在文档原位置覆盖保存'
  }
  Blockly.Blocks['saveWorkbook']=
    {
      init: function() {
        this.jsonInit(saveWorkbook);
      } 
    };
    python.pythonGenerator.forBlock['saveWorkbook'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('Workbook');
      var FILE =block.getFieldValue('FILE');
      var FILEPATH;
        var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        if(FILE!='')
        {
          FILEPATH = 'filename='+'\'' + FILE + '\'';
          code +=VAR+".save_excel_as(" + FILEPATH + "file_format=56)\n";
        }
        else
          {
            FILEPATH=FILE;
            code +=VAR+".save_excel(" + FILEPATH + ")\n";
          }
      // Return code.
      return code;
    }
}
function registerMoveActiveCell(){
  var MoveActiveCell ={
    "message0":"%1 MoveActiveCell(%2,%3)",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_number",
        "name":"row_change",
        "check":"number",

      },
      {
        "type": "field_number",
        "check":"number",
        "name":"column_change",
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":'{workbook} Set Active Cell {row} {column} : 设置活跃单元格\
    \nworkbook: Excel 文档变量名\
    \nrow: 行号\
    \ncolumn: 列号'
  }
  Blockly.Blocks['MoveActiveCell']=
    {
      init: function() {
        this.jsonInit(MoveActiveCell);
      } 
    };
    python.pythonGenerator.forBlock['MoveActiveCell'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('Workbook');
      var number1= block.getFieldValue('row_change');
        number1='row_change='+number1;
      var number2=block.getFieldValue('column_change');
        number2='column_change='+number2;
        var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        code +=VAR+".move_active_cell("+number1+","+number2+")\n";
      return code;
    }   
}
function registerSetActiveCell(){
  var SetActiveCell ={
    "message0":"%1 SetActiveCell(%2,%3)",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_number",
        "name":"row",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"column",
        "value": 100,
        "min":1,
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":"{workbook} Set Active Cell {row} {column} : 设置活跃单元格\
    \nworkbook: Excel 文档变量名\
    \nrow: 行号\
    \ncolumn: 列号",
  }
  Blockly.Blocks['SetActiveCell']=
    {
      init: function() {
        this.jsonInit(SetActiveCell);
      } 
    };
    python.pythonGenerator.forBlock['SetActiveCell'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('Workbook');
      var number1= block.getFieldValue('row');
      number1='row='+number1;
      var number2=block.getFieldValue('column');
      number2='column='+number2;
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="";
      for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=VAR+".set_active_cell("+number1+","+number2+")\n";
      return code;
    }   
}
function registerFetchCell(){
  var fetchCell ={
    "message0":"%1 fetch Cell (%2,%3) As %4",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_number",
        "name":"row",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"column",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "name": "VAR",
        "variable": "item",
        "variableTypes": [""],
        "check":"String"
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":"{workbook} Fetch Cell {row} {column} As {var} : 获取单元格\
    \nworkbook: Excel 文档变量名\
    \nrow: 行号，为空则采用当前获取行\
    \ncolumn: 列号，为空则采用当前活跃列\
    \nvar: 表示获取结果的变量名"
  }
  Blockly.Blocks['fetchCell']=
    {
      init: function() {
        this.jsonInit(fetchCell);
      } 
    };
    python.pythonGenerator.forBlock['fetchCell'] = function(block, generator) {
      // Collect argument strings.
      const Workbook = block.getFieldValue('Workbook');
      const VAR = block.getFieldValue('VAR');
      var number1= block.getFieldValue('row');
      number1='row='+number1;
      var number2=block.getFieldValue('column');
      number2='column='+number2;
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=VAR+"="+Workbook+".read_form_cells("+number1+","+number2+")\n";
      for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +="str"+VAR +"if" +VAR+"is not None else None\n";
      return code;
    }   
}
function registerFetchRow(){
  var fetchRow ={
    "message0":"%1 Fetch Row %2 (%3,%4) %5 As %6",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_number",
        "name":"row",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"columnF",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"columnT",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"header_row",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String"
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
  }
  Blockly.Blocks['fetchRow']=
    {
      init: function() {
        this.jsonInit(fetchRow);
      } 
    };
    python.pythonGenerator.forBlock['fetchRow'] = function(block, generator) {
      // Collect argument strings.
      const Workbook = block.getFieldValue('Workbook');
      var number1= block.getFieldValue('row');
      number1='row='+number1;
      var number2=block.getFieldValue('columnF');
      number2='column_from='+number2;
      var number3=block.getFieldValue('columnT');
      number3='column_to='+number3;
      var number4=block.getFieldValue('header_row');
      number4='header_row='+number4;
      var VAR=block.getFieldValue('VAR');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code += VAR +"="+Workbook+".read_row("+number1+","+number2+","+number3+","+number4+")\n";
      return code;
    }   
}
function registerFetchCol(){
  var fetchCol ={
    "message0":"%1 Fetch Column %2 (%3,%4) As %5",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_number",
        "name":"col",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"rowF",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"rowT",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "name": "VAR",
        "variable": "item",
        "variableTypes": [""],
        "check":"String"
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
  }
  Blockly.Blocks['fetchCol']=
    {
      init: function() {
        this.jsonInit(fetchCol);
      } 
    };
    python.pythonGenerator.forBlock['fetchCol'] = function(block, generator) {
      // Collect argument strings.
      const Workbook = block.getFieldValue('Workbook');
      var number1= block.getFieldValue('col');
      number1='column='+number1;
      var number2=block.getFieldValue('rowF');
      number2='row_from='+number2;
      var number3=block.getFieldValue('rowT');
      number3='row_to='+number3;
      var VAR=block.getFieldValue('VAR');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=VAR+"="+Workbook+".read_col("+number1+","+number2+","+number3+")\n";
      return code;
    }   

}
function registerFetchArea(){
  var fetchArea ={
    "message0":"%1 Fetch Area (%2,%3), (%4,%5)  %6 As %7",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"rowF",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"rowT",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"colF",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"colT",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_dropdown",
        "name": "with_header",
        "options": [
          [ "false", "false" ],
          ["true","true"]
        ]
      },
      {
        "type": "field_input",
        "name": "VAR",
        "variable": "item",
        "variableTypes": [""],
        "check":"String"
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":"{workbook} Fetch area {row_from} {row_to} {column_from} {column_to} {header} As {var} : 获取一个区域\
    \nworkbook: Excel 文档变量名\
    \nrow_from: 起点行号\
    \nrow_to: 终点行号\
    \ncolumn_from: 起点列号\
    \ncolumn_to: 终点列号\
    \nheader: 是否将第一行作为列名\
    \nvar: 表示获取结果的变量名"
  }
  Blockly.Blocks['fetchArea']=
    {
      init: function() {
        this.jsonInit(fetchArea);
      } 
    };
    python.pythonGenerator.forBlock['fetchArea'] = function(block, generator) {
      // Collect argument strings.
      const Workbook = block.getFieldValue('Workbook');
      var number1= block.getFieldValue('rowF');
      number1='row_from='+number1;
      var number2=block.getFieldValue('rowT');
      number2='row_to='+number2;
      var number3=block.getFieldValue('colT');
      number3='column_from='+number3;
      var number4=block.getFieldValue('colF');
      number4='column_to='+number4;
      var header=block.getFieldValue('with_header');
      header='header='+header;
      var VAR=block.getFieldValue('VAR');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=VAR+"="+Workbook+".read_area("+number1+","+number2+","+number3+","+number4+","+header+")\n";
      return code;
    }   

}
function registerInsertCol(){
  var InsertCol ={
    "message0":"%1 Insert Col to %2 %3",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"column",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"string",
        "name":"col_content",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
  }
  Blockly.Blocks['InsertCol']=
    {
      init: function() {
        this.jsonInit(InsertCol);
      } 
    };
    python.pythonGenerator.forBlock['InsertCol'] = function(block, generator) {
      // Collect argument strings.
      const workbook = block.getFieldValue('Workbook');
      const column= block.getFieldValue('column');
      const col_content=block.getFieldValue('col_content');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        code+=workbook+".insert_column(column="+column+","+" column_content="+col_content+")\n";
      return code;
    }   

}
function registerInsertRow(){
  var InsertRow ={
    "message0":"%1 Insert Row to %2 %3 %4",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"row",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"string",
        "name":"row_content",
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"header",
        "value": 100,
        "min":1,
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
  }
  Blockly.Blocks['InsertRow']=
    {
      init: function() {
        this.jsonInit(InsertRow);
      } 
    };
    python.pythonGenerator.forBlock['InsertRow'] = function(block, generator) {
      // Collect argument strings.
      const workbook = block.getFieldValue('Workbook');
      const row= block.getFieldValue('row');
      const row_content=block.getFieldValue('row_content');
      const header_row=block.getFieldValue('header');
      var code ="";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code+=workbook+".insert_row(row="+row+","+"row_content="+row_content+",header_row="+header_row+")\n";
      return code;
    }   

}
function registerSetCellValue(){
  var SetCellValue ={
    "message0":"%1 SetCellValue(%2,%3) %4",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_number",
        "name":"row",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "check":"number",
        "name":"column",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_number",
        "name":"value",
        "check":"number",
        "value": 100,
        "min":1,
      }
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":"{workbook} Set Cell {row} {column} {value} : 设置单元格的值\
    workbook: Excel 文档变量名\
    row: 行号，为空则采用当前获取行\
    column: 列号，为空则采用当前活跃列\
    value: 待写入的值"
    
  }
  Blockly.Blocks['SetCellValue']=
    {
      init: function() {
        this.jsonInit(SetCellValue);
      } 
    };
    python.pythonGenerator.forBlock['SetCellValue'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('Workbook');
      var number1= block.getFieldValue('row');
      number1='row='+number1;
      var number2=block.getFieldValue('column');
      number2='column='+number2;
      var value=block.getFieldValue('value');
      value='value='+value;
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code +=VAR+".set_active_cell("+number1+","+number2+","+value+"number_format='@'"+")\n";
      return code;
    }   
}
function registerSettoBlock(){
  var setto ={
    "message0":"Set %1 to %2",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String"
      },
      {
        "type": "field_input",
        "name": "exp",
        "check":"String"
      }],
    "previousStatement": null,
    "nextStatement": null,
    "colour":200,
    "tooltip":'{workbook} Save Workbook {path} : 保存 Excel 文档 \
    \nworkbook: Excel 文档变量名 \
    \npath: 目标保存路径，为空表示在文档原位置覆盖保存'
  }
  Blockly.Blocks['setto']=
    {
      init: function() {
        this.jsonInit(setto);
      } 
    };
    python.pythonGenerator.forBlock['setto'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
      var exp =block.getFieldValue('exp');
        var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        code +=VAR+"="+exp;
      // Return code.
      return code;
    }
}
function registerCreateSheet(){
  var CreateSheet ={
    "message0":"%1 CreateSheet %2",
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
    
  }
  Blockly.Blocks['CreateSheet']=
    {
      init: function() {
        this.jsonInit(CreateSheet);
      } 
    };
    python.pythonGenerator.forBlock['CreateSheet'] = function(block, generator) {
      const Workbook = block.getFieldValue('Workbook');
      const name= block.getFieldValue('name');
      var code ="";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code +=Workbook+".add_new_sheet("+name+")\n";
      return code;
    }   
}
function registerSetActiveSheet(){
  var SetActiveSheet ={
    "message0":"%1 SetActiveSheet %2",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "name": "name",
        "check":"String"
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":240,
    
  }
  Blockly.Blocks['SetActiveSheet']=
    {
      init: function() {
        this.jsonInit(SetActiveSheet);
      } 
    };
    python.pythonGenerator.forBlock['SetActiveSheet'] = function(block, generator) {
      const Workbook = block.getFieldValue('Workbook');
      const name =block.getFieldValue('name');
      var code ="";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code +=Workbook+".set_active_worksheet("+name+")\n";
      return code;
    }   
}
function registerMergeSheet(){
  var MergeSheet ={
    "message0":"%1 MergeSheet %2",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "name": "name",
        "check":"String"
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":240,
    
  }
  Blockly.Blocks['MergeSheet']=
    {
      init: function() {
        this.jsonInit(MergeSheet);
      } 
    };
    python.pythonGenerator.forBlock['MergeSheet'] = function(block, generator) {
      // Collect argument strings.
      const Workbook = block.getFieldValue('Workbook');
      const name =block.getFieldValue('name');
      var code ="";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code +=Workbook+".set_active_worksheet("+name+")\n";
      return code;
    }   
}
function registerCompareBlock(){{
  var Compare ={
    "type":"condition",
    "message0":"Compare %1 %2 %3 %4 %5",
    "args0": [
      {
        "type": "field_dropdown",
        "name": "value_type",
        "options": [
          [ "int", "int" ],
          [ "float", "float" ],
          ["str","string"],
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
          [ "<", "<" ],
          [ "<=", "<=" ],
          ["==","=="],
          [">=",">="],
          [">",">"],
        ]
      },
      {
        "type": "field_input",
        "name":"exp2",
        "check":"string",
      },
      {
        "type": "input_value",
        "name": "prolong",
      }
    ],
    "colour":220,
    "output": null,
  }
  Blockly.Blocks['Compare']=
    {
      init: function() {
        this.jsonInit(Compare);
      } 
    };
    python.pythonGenerator.forBlock['Compare'] = function(block, generator) {
      // Collect argument strings.
      var valueType = block.getFieldValue('value_type');
      var exp1 =block.getFieldValue('exp1');
      var comparation =block.getFieldValue('comparation');
      var exp2 =block.getFieldValue('exp2');
      var prolong=generator.statementToCode(block,'prolong');
      var code =valueType+" ("+exp1+") "+comparation+" "+valueType+" ("+exp2+")"+prolong;
      return code;
    }   

}}
function registerIfBlock(){{
  var IF ={
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
  }
  Blockly.Blocks['IF']=
    {
      init: function() {
        this.jsonInit(IF);
      } 
    };
    python.pythonGenerator.forBlock['IF'] = function(block, generator) {
      // Collect argument strings.
      
      var condition =generator.statementToCode(block,'condition');
      depth+=1;
      var content =generator.statementToCode(block,'content');
      depth-=1;
      var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +="if"+condition+" :\n";
      code +=content;
      
      return code;
    }   

}}
function registerElifBlock(){{
  var Elif ={
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
  }
  Blockly.Blocks['Elif']=
    {
      init: function() {
        this.jsonInit(Elif);
      } 
    };
    python.pythonGenerator.forBlock['Elif'] = function(block, generator) {
      // Collect argument strings.
      var condition =generator.statementToCode(block,'condition');
      depth+=1;
      var content =generator.statementToCode(block,'content');
      depth-=1;
      var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +="elif"+condition+" :\n";
      code +=content;
      return code;
    }   

}}
function registerElseBlock(){{
  var Else ={
    "message0":"Else %1",
    "args0": [
      {
        "type": "input_statement",
        "name":"content",
        "check":"string",
      },
    ],
    "previousStatement": null,
    "colour":220,
  }
  Blockly.Blocks['Else']=
    {
      init: function() {
        this.jsonInit(Else);
      } 
    };
    python.pythonGenerator.forBlock['Else'] = function(block, generator) {
      // Collect argument strings.
      depth+=1;
      var content =generator.statementToCode(block,'content');
      depth-=1;
      var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +="else :\n";
      code +=content;
      return code;
    }   

}}
function registerForBlock(){{
  var For ={
    "message0":"for %1 from %2 to %3 : %4",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String"
      },
      {
        "type": "field_number",
        "name":"start",
        "check":"number",

      },
      {
        "type": "field_number",
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
    "colour":220,
  }
  Blockly.Blocks['For']=
    {
      init: function() {
        this.jsonInit(For);
      } 
    };
    python.pythonGenerator.forBlock['For'] = function(block, generator) {
      // Collect argument strings.
      var VAR = block.getFieldValue('VAR');
      var start = block.getFieldValue('start');
      var end = block.getFieldValue('end');
      depth+=1;
      var content =generator.statementToCode(block,'content');
      depth-=1;
      var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +="for "+VAR+"in range"+"("+start+","+end+"):\n";
      code+=content;
      return code;
    }   

}}
function registerForeachBlock(){{
  var Foreach ={
    "message0":"for each %1  in %2 :%3",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String"
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
    "colour":220,
  }
  Blockly.Blocks['Foreach']=
    {
      init: function() {
        this.jsonInit(Foreach);
      } 
    };
    python.pythonGenerator.forBlock['Foreach'] = function(block, generator) {
      // Collect argument strings.
      var VAR = block.getFieldValue('VAR');
      var iterableVar = block.getFieldValue('iterable_var');
      depth+=1;
      var content =generator.statementToCode(block,'content');
      depth-=1;
      var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +="for "+VAR+"in"+iterableVar+":\n";
      code+=condition;
      return code;
    }   

}}
function registerAndBlock(){{
  var and ={
    "message0":"and %1",
    "args0": [
      {
        "type": "input_value",
        "name":"condition1",
      },
    ],
    "output": null,
    "colour":220,
  }
  Blockly.Blocks['and']=
    {
      init: function() {
        this.jsonInit(and);
      } 
    };
    python.pythonGenerator.forBlock['and'] = function(block, generator) {
      // Collect argument strings.
      var iterableVar = block.getFieldValue('iterable_var');
      var condition1 =generator.statementToCode(block,'condition1');
      var code='';
      code +="and"+condition1;
      return code;
    }   

}}

function registerOrBlock(){{
  var or ={
    "message0":"or %1",
    "args0": [
      {
        "type": "input_value",
        "name":"condition1",
      },
    ],
    "output": null,
    "colour":220,
  }
  Blockly.Blocks['or']=
    {
      init: function() {
        this.jsonInit(or);
      } 
    };
    python.pythonGenerator.forBlock['or'] = function(block, generator) {
      // Collect argument strings.
      var condition1 =generator.statementToCode(block,'condition1');
      var code='';
      code +="or"+condition1;
      return code;
    }   

}}
function registerNotBlock(){{
  var not ={
    "message0":"not %1",
    "args0": [
      {
        "type": "input_value",
        "name":"condition1",
      },
    ],
    "colour":220,
    "output":null,
  }
  Blockly.Blocks['not']=
    {
      init: function() {
        this.jsonInit(not);
      } 
    };
    python.pythonGenerator.forBlock['not'] = function(block, generator) {
      // Collect argument strings.
      var iterableVar = block.getFieldValue('iterable_var');
      var condition1 =generator.statementToCode(block,'condition1');
      var code='';
      code +="not"+condition1;
      return code;
    }   

}}
function registerHelpOption() {
  const helpItem = {
    displayText: 'Help! There are no blocks',
    // Use the workspace scope in the precondition function to check for blocks on the workspace.
    preconditionFn: function(scope) {
      if (!scope.workspace.getTopBlocks().length) {
        return 'enabled';
      }
      return 'hidden';
    },
    // Use the workspace scope in the callback function to add a block to the workspace.
    callback: function(scope) {
      Blockly.serialization.blocks.append({
        'type': 'text',
        'fields': {
          'TEXT': 'Now there is a block'
        }
      }, scope.workspace);
    },
    scopeType: Blockly.ContextMenuRegistry.ScopeType.WORKSPACE,
    id: 'help_no_blocks',
    weight: 100,
  };
  Blockly.ContextMenuRegistry.registry.register(helpItem);
}

function registerOutputOption() {
  const outputOption = {
    displayText: 'I have an output connection',
    // Use the block scope in the precondition function to hide the option on blocks with no
    // output connection.
    preconditionFn: function(scope) {
      if (scope.block.outputConnection) {
        return 'enabled';
      }
      return 'hidden';
    },
    callback: function (scope) {
    },
    scopeType: Blockly.ContextMenuRegistry.ScopeType.BLOCK,
    id: 'block_has_output',
    // Use a larger weight to push the option lower on the context menu.
    weight: 200,
  };
  Blockly.ContextMenuRegistry.registry.register(outputOption);
}

function registerDisplayOption() {
  const displayOption = {
    // Use the block scope to set display text dynamically based on the type of the block.
    displayText: function(scope) {
      if (scope.block.type.startsWith('text')) {
        return 'Text block';
      } else if (scope.block.type.startsWith('controls')) {
        return 'Controls block';
      } else {
        return 'Some other block';
      }
    },
    preconditionFn: function (scope) {
      return 'enabled';
    },
    callback: function (scope) {
    },
    scopeType: Blockly.ContextMenuRegistry.ScopeType.BLOCK,
    id: 'display_text_example',
    weight: 100,
  };
  Blockly.ContextMenuRegistry.registry.register(displayOption);
}
