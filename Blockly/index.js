'use strict';
let workspace = null;

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
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="\n\t"+VAR+"=ExcelApplication()";
      code +="\n\t"+VAR+".open_application(visible=True)";
        code +="\n\t"+VAR+".open_workbook("+FILEPATH+")";
      return code;
    }       
          
}
function registerSaveWorkbook(){
  var saveWorkbook ={
    "message0":" and Save Workbook %1",
    "args0": [
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
      const VAR = block.getRootBlock().getFieldValue('VAR');
      var FILE =block.getFieldValue('FILE');
      var FILEPATH;
      if(FILE!='')
        FILEPATH = 'filename='+'\'' + FILE + '\'';
      else
        FILEPATH=FILE;
      var code ="\n\t"+VAR+".save_workbook("+FILEPATH+")";
      // Return code.
      return code;
    }
}
function registerMoveActiveCell(){
  var MoveActiveCell ={
    "message0":"MoveActiveCell(%1,%2)",
    "args0": [
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
      const VAR = block.getRootBlock().getFieldValue('VAR');
      var number1= block.getFieldValue('row_change');
        number1='row_change='+number1;
      var number2=block.getFieldValue('column_change');
        number2='column_change='+number2;
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="\n\t"+VAR+".move_active_cell("+number1+","+number2+")";
      return code;
    }   
}
function registerSetActiveCell(){
  var SetActiveCell ={
    "message0":"SetActiveCell(%1,%2)",
    "args0": [
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
    
  }
  Blockly.Blocks['SetActiveCell']=
    {
      init: function() {
        this.jsonInit(SetActiveCell);
      } 
    };
    python.pythonGenerator.forBlock['SetActiveCell'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getRootBlock().getFieldValue('VAR');
      var number1= block.getFieldValue('row');
      number1='row='+number1;
      var number2=block.getFieldValue('column');
      number2='column='+number2;
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="\n\t"+VAR+".set_active_cell("+number1+","+number2+")";
      return code;
    }   
}
function registerFetchCell(){
  var fetchCell ={
    "message0":"fetch Cell (%1,%2) As %3",
    "args0": [
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
    
  }
  Blockly.Blocks['fetchCell']=
    {
      init: function() {
        this.jsonInit(fetchCell);
      } 
    };
    python.pythonGenerator.forBlock['fetchCell'] = function(block, generator) {
      // Collect argument strings.
      const Workbook = block.getRootBlock().getFieldValue('VAR');
      const VAR = block.getFieldValue('VAR');
      var number1= block.getFieldValue('row');
      number1='row='+number1;
      var number2=block.getFieldValue('column');
      number2='column='+number2;
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="\n\t"+VAR+"="+Workbook+".read_form_cells("+number1+","+number2+")";
      return code;
    }   
}
function registerFetchRow(){
  var fetchRow ={
    "message0":"Fetch Row %1 (%2,%3) As %4",
    "args0": [
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
    
  }
  Blockly.Blocks['fetchRow']=
    {
      init: function() {
        this.jsonInit(fetchRow);
      } 
    };
    python.pythonGenerator.forBlock['fetchRow'] = function(block, generator) {
      // Collect argument strings.
      const Workbook = block.getRootBlock().getFieldValue('VAR');
      var number1= block.getFieldValue('row');
      number1='row='+number1;
      var number2=block.getFieldValue('columnF');
      number2='column_from='+number2;
      var number3=block.getFieldValue('columnT');
      number3='column_to='+number3;
      var VAR=block.getFieldValue('VAR');
      var code ="\n\t"+VAR+"="+Workbook+".read_row("+number1+","+number2+","+number3+")";
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      return code;
    }   
}
function registerFetchCol(){
  var fetchCol ={
    "message0":"Fetch Column %1 (%2,%3) As %4",
    "args0": [
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
  }
  Blockly.Blocks['fetchCol']=
    {
      init: function() {
        this.jsonInit(fetchCol);
      } 
    };
    python.pythonGenerator.forBlock['fetchCol'] = function(block, generator) {
      // Collect argument strings.
      const Workbook = block.getRootBlock().getFieldValue('VAR');
      var number1= block.getFieldValue('col');
      number1='column='+number1;
      var number2=block.getFieldValue('rowF');
      number2='row_from='+number2;
      var number3=block.getFieldValue('rowT');
      number3='row_to='+number3;
      var VAR=block.getFieldValue('VAR');
      var code ="\n\t"+VAR+"="+Workbook+".read_col("+number1+","+number2+","+number3+")";
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      return code;
    }   

}
function registerFetchArea(){
  var fetchArea ={
    "message0":"Fetch Area (%1,%2), (%3,%4)  %5 As %6",
    "args0": [
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
        "type": "field_number",
        "check":"number",
        "name":"colT",
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
  }
  Blockly.Blocks['fetchArea']=
    {
      init: function() {
        this.jsonInit(fetchArea);
      } 
    };
    python.pythonGenerator.forBlock['fetchArea'] = function(block, generator) {
      // Collect argument strings.
      const Workbook = block.getRootBlock().getFieldValue('VAR');
      var number1= block.getFieldValue('rowF');
      number1='row_from='+number1;
      var number2=block.getFieldValue('rowT');
      number2='row_to='+number2;
      var number3=block.getFieldValue('colT');
      number3='column_from='+number3;
      var number4=block.getFieldValue('colF');
      number4='column_to='+number4;
      var header=block.getFieldValue('header');
      header='header='+header;
      var VAR=block.getFieldValue('VAR');
      var code ="\n\t"+VAR+"="+Workbook+".read_area("+number1+","+number2+","+number3+","+number4+","+header+")";
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      return code;
    }   

}
function registerInsertCol(){
  var InsertCol ={
    "message0":"Insert Col to %1",
    "args0": [
      {
        "type": "field_number",
        "check":"number",
        "name":"N1",
        "value": 100,
        "min":1,
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
      const VAR = block.getRootBlock().getFieldValue('VAR');
      const number1= block.getFieldValue('N1');
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
     var code ="\n\tfor i in range("+number1+","+number2+"):";
     code +="\n\t\tr = "+VAR+".read_col(header=True)";
     code +="\n\t\tprint(r)";
     code +="\n\t\tapp.move_active_cell(1, 0)"; 
      return code;
    }   

}
function registerInsertRow(){
  var InsertRow ={
    "message0":"Insert Row to %1",
    "args0": [
      {
        "type": "field_number",
        "check":"number",
        "name":"N1",
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
      const VAR = block.getRootBlock().getFieldValue('VAR');
      const number1= block.getFieldValue('N1');
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
     var code ="\n\tfor i in range("+number1+","+number2+"):";
     code +="\n\t\tr = "+VAR+".read_row(header=True)";
     code +="\n\t\tprint(r)";
     code +="\n\t\tapp.move_active_cell(1, 0)"; 
      return code;
    }   

}
function registerSetCellValue(){
  var SetCellValue ={
    "message0":"SetCellValue(%1,%2)%3",
    "args0": [
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
    
  }
  Blockly.Blocks['SetCellValue']=
    {
      init: function() {
        this.jsonInit(SetCellValue);
      } 
    };
    python.pythonGenerator.forBlock['SetCellValue'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getRootBlock().getFieldValue('VAR');
      var number1= block.getFieldValue('row');
      number1='row='+number1;
      var number2=block.getFieldValue('column');
      number2='column='+number2;
      var value=block.getFieldValue('value');
      value='value='+value;
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="\n\t"+VAR+".set_active_cell("+number1+","+number2+","+value+")";
      return code;
    }   
}
function registerCreateSheet(){
  var CreateSheet ={
    "message0":"CreateSheet %1",
    "args0": [
      {
        "type": "field_input",
        "name": "FILE",
        "check":"String"
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
      // Collect argument strings.
      const VAR = block.getRootBlock().getFieldValue('VAR');
      const FILEPATH = '\'' + block.getFieldValue('FILE') + '\'';
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="\n\t"+VAR+".open_workbook("+FILEPATH+")";
      return code;
    }   
}
function registerSetActiveSheet(){
  var SetActiveSheet ={
    "message0":"SetActiveSheet %1",
    "args0": [
      {
        "type": "field_input",
        "name": "FILE",
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
      // Collect argument strings.
      const VAR = block.getRootBlock().getFieldValue('VAR');
      const number1= block.getFieldValue('N1');
      const number2=block.getFieldValue('N2');
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="\n\t"+VAR+".set_active_cell("+number1+","+number2+")";
      return code;
    }   
}
function registerMergeSheet(){
  var MergeSheet ={
    "message0":"MergeSheet %1",
    "args0": [
      {
        "type": "field_input",
        "name": "FILE",

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
      const VAR = block.getRootBlock().getFieldValue('VAR');
      const FILEPATH = '\'' + block.getFieldValue('FILE') + '\'';
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="\n\t"+VAR+".set_active_cell("+number1+","+number2+")";
      return code;
    }   
}
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
