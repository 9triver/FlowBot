'use strict';
let workspace = null;
var depth=1;
var logicOperator =" null ";

const fs=require('fs');
const exec = require('child_process').exec;
const path = require('node:path');
const { app, BrowserWindow,shell,ipcMain} = require('electron');
const { dialog } = require('electron');
const createWindow = () => {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  })
  win.loadFile('index.html')
// 在主进程中.
win.webContents.session.on('will-download', (event, item, webContents) => {
  // 无需对话框提示， 直接将文件保存到路径
  if(item.getFilename()=='tasks.py'){
  item.setSavePath(__dirname+"\\RPA\\test\\tasks.py")
  item.on('updated', (event, state) => {
    if (state === 'interrupted') {
      console.log('Download is interrupted but can be resumed')
    } else if (state === 'progressing') {
      if (item.isPaused()) {
        console.log('Download is paused')
      } else {
        console.log(`Received bytes: ${item.getReceivedBytes()}`)
      }
    }
  })
  item.once('done', (event, state) => {
    if (state === 'completed') {
      console.log('Download successfully')
    } else {
      console.log(`Download failed: ${state}`)
    }
  })
  let myPath = "//RPA//test"
  let cmdStr1 = 'rcc.exe run';
  let cmdPath = __dirname+myPath
  // 子进程名称
  let workerProcess
  runExec(cmdStr1)
  function runExec (cmdStr) {
    workerProcess = exec(cmdStr, { cwd: cmdPath })
    // 打印正常的后台可执行程序输出
    workerProcess.stdout.on('data', function (data) {
      console.log('stdout: ' + data)
    })
    // 打印错误的后台可执行程序输出
    workerProcess.stderr.on('data', function (data) {
      console.log('stderr: ' + data)
    })
    // 退出之后的输出
  }
  }
  else if(item.getFilename()=='myBlock.json')
  {
    
  }
})
}
async function handleFileOpen () {
  const { canceled, filePaths } = await dialog.showOpenDialog({})
  if (!canceled) {
    let file=fs.readFileSync(filePaths[0]);
    let ob= JSON.parse(file);
    //let cur =JSON.stringify(ob);
    return ob
}
}
async function handleFileOpenPath () {
  const { canceled, filePaths } = await dialog.showOpenDialog({})
  if (!canceled) {
    //let cur =JSON.stringify(ob);
    return filePaths[0]
}
}
app.whenReady().then(() => {
  ipcMain.handle('dialog:openFile',handleFileOpen);
  ipcMain.handle('dialog:openFilePath',handleFileOpenPath);
  createWindow();
  // shell.openPath(".\\tasks.py")
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})









function start() {
  registerFirstContextMenuOptions();

 
  registerOpenWorkbook();
  registerAddWorkbook();
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
         {spacing: 30,
          length: 7,
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
    "message0": "打开指定目录(%1)下的Workbook，将其命名为%2",
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
    "message0":"将名为%1的Workbook保存到路径%2",
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
          code +=VAR+".save_excel_as(" + FILEPATH + ",file_format=56)\n";
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
function registerAddWorkbook(){
  var addWorkbook ={
    "message0":"新增一个Workbook，将其命名为%1",
    "args0": [
      {
        "type": "field_input",
        "name": "VAR",
        "check":"String"
      }],
    "previousStatement": null,
    "nextStatement": null,
    "colour":200,
    "tooltip":'Add Workbook as {var} : 保存 Excel 文档 \
    \nworkbook: Excel 文档变量名 \
    \npath: 目标保存路径，为空表示在文档原位置覆盖保存'
  }
  Blockly.Blocks['addWorkbook']=
    {
      init: function() {
        this.jsonInit(addWorkbook);
      } 
    };
    python.pythonGenerator.forBlock['addWorkbook'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
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
        code +=VAR+".add_new_workbook()\n";
      return code;
    }
}
function registerMoveActiveCell(){
  var MoveActiveCell ={
    "message0":"移动名为%1的Workbook中的\n活跃单元格到(第%2行,第%3列)",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "name":"row_change",
        "check":"String",

      },
      {
        "type": "field_input",
        "check":"string",
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
    "message0":"设置名为%1的Workbook中的\n活跃单元格到坐标(第%2行,第%3列)",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "name":"row",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
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
    "message0":"获取名为%1的Workbook中的\n(第%2行,第%3列)单元格，将其命名为%4",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "name":"row",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
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
      if(number1!='')
      number1='row='+number1;
      var number2=block.getFieldValue('column');
      if(number2!='')
      number2='column='+number2;
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        if(number1==''&&number2!='')
        {
          code +=VAR+"="+Workbook+".read_form_cells("+number2+")\n";
        }
        else if(number2==''&&number1!='')
        {
          code +=VAR+"="+Workbook+".read_form_cells("+number1+")\n";
        }
        else if(number1!=''&&number2!='')
          code +=VAR+"="+Workbook+".read_form_cells("+number1+","+number2+")\n";
        else
        code +=VAR+"="+Workbook+".read_form_cells()\n";
      for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +="str("+VAR +")if " +VAR+" is not None else None\n";
      return code;
    }   
}
function registerFetchRow(){
  var fetchRow ={
    "message0":"获取名为%1的Workbook中的\n第%2行(从第%3列到第%4列)(头部行:%5)，将该变量命名为 %6",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "name":"row",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"columnF",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"columnT",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
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
      if(number1!='')
      number1='row='+number1;
      var number2=block.getFieldValue('columnF');
      if(number2!='')
      number2='column_from='+number2;
      var number3=block.getFieldValue('columnT');
      if(number3!='')
      number3='column_to='+number3;
      var number4=block.getFieldValue('header_row');
      if(number4!='')
      number4='header_row='+number4;
      var VAR=block.getFieldValue('VAR');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      if(number4=='')
      { 
        if(number3=='')
        {
          if(number2=='')
            {
            if(number1=='')
            code += VAR +"="+Workbook+".read_row()\n"; 
            else
            code += VAR +"="+Workbook+".read_row("+number1+")\n";
            }
          else if(number1=='')
          {
            code += VAR +"="+Workbook+".read_row("+number2+")\n";
          }
          else
            code += VAR +"="+Workbook+".read_row("+number1+","+number2+")\n";
        }
        else if(number2=='')
        { 
          if(number1=='')
            {
              code += VAR +"="+Workbook+".read_row("+number3+")\n";
            }
          else 
          code += VAR +"="+Workbook+".read_row("+number1+","+number3+")\n";
        }
        else if(number1=='')
        code += VAR +"="+Workbook+".read_row("+number2+","+number3+")\n";
        else
        code += VAR +"="+Workbook+".read_row("+number1+","+number2+","+number3+")\n";
      }
      else if(number3=='')
      {
        if(number2=='')
        {
          if(number1=='')
          {
            code += VAR +"="+Workbook+".read_row("+number4+")\n";
          }
          else
          code += VAR +"="+Workbook+".read_row("+number1+","+number4+")\n";
        }
        else if(number1=='')
        {
          code += VAR +"="+Workbook+".read_row("+number3+")\n";
        }
        else
        code += VAR +"="+Workbook+".read_row("+number1+","+number2+","+number4+")\n";
      }
      else if(number2=='')
      { 
        if(number1=='')
        {
          code += VAR +"="+Workbook+".read_row("+number3+","+number4+")\n";
        }
        else
        code += VAR +"="+Workbook+".read_row("+number1+","+number3+","+number4+")\n";
      }
      else if(number1=='')
      { 
        code += VAR +"="+Workbook+".read_row("+number2+","+number3+","+number4+")\n";
      }
      else
      code += VAR +"="+Workbook+".read_row("+number1+","+number2+","+number3+","+number4+")\n";
      return code;
    }   
}
function registerFetchCol(){
  var fetchCol ={
    "message0":"获取名为%1的Workbook中的\n第%2列(从哪%3行到哪%4行)，将该变量命名为 %5",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "name":"col",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"rowF",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
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
      if(number1!='')
      number1='column='+number1;
      var number2=block.getFieldValue('rowF');
      if(number2!='')
      number2='row_from='+number2;
      var number3=block.getFieldValue('rowT');
      if(number3!='')
      number3='row_to='+number3;
      var VAR=block.getFieldValue('VAR');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      if(number3=='')
        {
          if(number2=='')
          {
            if(number1=='')
            {
              code +=VAR+"="+Workbook+".read_col()\n";
            }
            else
            code +=VAR+"="+Workbook+".read_col("+number1+")\n";
          }
          else if(number1=='')
          {
            code +=VAR+"="+Workbook+".read_col("+number2+")\n";
          }
          else
          code +=VAR+"="+Workbook+".read_col("+number1+","+number2+")\n";
        }
      else if(number2=='')
      {
        if(number1=='')
        {
          code +=VAR+"="+Workbook+".read_col("+number3+")\n";
        }
        else 
        code +=VAR+"="+Workbook+".read_col("+number1+","+number3+")\n";
      }
      else if(number1=='')
      {
        code +=VAR+"="+Workbook+".read_col("+number2+","+number3+")\n";
      }
      else
        code +=VAR+"="+Workbook+".read_col("+number1+","+number2+","+number3+")\n";
      return code;
    }   

}
function registerFetchArea(){
  var fetchArea ={
    "message0":"获取名为%1的Workbook的区域\n(第%2行到第%3行)(第%4列到第%5列)(头部:%6)，将其变量命名为%7",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"rowF",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"rowT",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"colF",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"colT",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_dropdown",
        "name": "with_header",
        "options": [
          [ "False", "False" ],
          ["True","True"]
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
      if(header!='')
      header='header='+header;
      var VAR=block.getFieldValue('VAR');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      if(header=='')
      code +=VAR+"="+Workbook+".read_area("+number1+","+number2+","+number3+","+number4+")\n";
      else
      code +=VAR+"="+Workbook+".read_area("+number1+","+number2+","+number3+","+number4+","+header+")\n";
      return code;
    }   

}
function registerInsertCol(){
  var InsertCol ={
    "message0":"往名为%1的Workbook之中插入新列(列号:%2,列值:%3)",
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
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
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
    "message0":"往名为%1的Workbook中插入新行(行号:%2,行值:%3)(行头:%4)",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"row",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "check":"string",
        "name":"row_content",
      },
      {
        "type": "field_input",
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
      if(header!='')
      code+=workbook+".insert_row(row="+row+","+"row_content="+row_content+",header_row="+header_row+")\n";
      else
      code+=workbook+".insert_row(row="+row+","+"row_content="+row_content+")\n";
      return code;
    }   

}
function registerSetCellValue(){
  var SetCellValue ={
    "message0":"设置名为%1的Workbook的单元格(第%2行,第%3列)新值:%4",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String"
      },
      {
        "type": "field_input",
        "name":"row",
        "check":"number",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"column",
        "value": 100,
        "min":1,
      },
      {
        "type": "field_input",
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
      if(number1!='')
      number1='row='+number1;
      var number2=block.getFieldValue('column');
      if(number2!='')
      number2='column='+number2;
      var value=block.getFieldValue('value');
      value='value='+value;
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      
      if(number2=='')
        {
          if(number1=='')
          {
            code +=VAR+".set_active_cell("+value+","+"number_format='@'"+")\n";
          }
          else
          code +=VAR+".set_active_cell("+number1+","+value+","+"number_format='@'"+")\n";
        }
      else if(number1=='')
        {
          code +=VAR+".set_active_cell("+number2+","+value+","+"number_format='@'"+")\n";
        }
      code +=VAR+".set_active_cell("+number1+","+number2+","+value+","+"number_format='@'"+")\n";
      return code;
    }   
}
function registerSettoBlock(){
  var setto ={
    "message0":"将变量%1设置为%2",
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
        code +=VAR+"="+exp+"\n";
      // Return code.
      return code;
    }
}
function registerCreateSheet(){
  var CreateSheet ={
    "message0":"在%1Workbook里创建新的sheet %2",
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
      code +=Workbook+".add_new_sheet('"+name+"')\n";
      return code;
    }   
}
function registerSetActiveSheet(){
  var SetActiveSheet ={
    "message0":"在%1Workbook里设置活跃sheet%2",
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
    "message0":"在%1Workbook里合并%2和sheet",
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
    "message0":" %1 %2 %3 %4",
    "args0": [
      {
        "type": "field_dropdown",
        "name": "value_type",
        "options": [
          [ "int", "int" ],
          [ "float", "float" ],
          ["str","str"],
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
    ],
    "output":null,
    "previousStatement": null,
    "nextStatement":null,
    "colour":220,
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
      var nextBlock = block.getNextBlock();
      if(nextBlock!=null && logicOperator!=" null ")
      {
        //alert(nextBlock);
        var code =valueType+"("+exp1+") "+comparation+" "+valueType+"("+exp2+")"+logicOperator;
      }
      //var prolong=generator.statementToCode(block,'prolong');
      else 
        var code =valueType+"("+exp1+") "+comparation+" "+valueType+"("+exp2+")";
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
    "nextStatement": null,
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
      code +="for "+VAR+" in range"+"("+start+","+end+"):\n";
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
    "nextStatement": null,
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
  }
  Blockly.Blocks['and']=
    {
      init: function() {
        this.jsonInit(and);
      } 
    };
    python.pythonGenerator.forBlock['and'] = function(block, generator) {
      // Collect argument strings.
      var cur=logicOperator.toString();
      logicOperator=" and ";
      var condition1 =generator.statementToCode(block,'condition1');
      logicOperator=cur.toString();
      var code='';
      var NextBlock = block.getNextBlock();
      var previousBlock=block.getPreviousBlock();
      if(logicOperator!=" null "&&NextBlock!=null)
      code+="("+condition1+")"+logicOperator;
      else
      code+="("+condition1+")";
      return code;
    }   

}}

function registerOrBlock(){{
  var or ={
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
  }
  Blockly.Blocks['or']=
    {
      init: function() {
        this.jsonInit(or);
      } 
    };
    python.pythonGenerator.forBlock['or'] = function(block, generator) {
      // Collect argument strings.
      var cur=logicOperator.toString();
      logicOperator=" or ";
      var condition1 =generator.statementToCode(block,'condition1');
      logicOperator=cur.toString();
      var code='';
      var previousBlock = block.getPreviousBlock();
      var NextBlock = block.getNextBlock();
      if(logicOperator!=" null "&&NextBlock!=null)
      code+="("+condition1+")"+logicOperator;
      else
      code+="("+condition1+")";
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
    "output": null,
    "previousStatement":null,
    "nextStatement":null,
  }
  Blockly.Blocks['not']=
    {
      init: function() {
        this.jsonInit(not);
      } 
    };
    python.pythonGenerator.forBlock['not'] = function(block, generator) {
      // Collect argument strings.
      var condition1 =generator.statementToCode(block,'condition1');
      var code='';
      var nextBlock =block.getNextBlock();
      if(logicOperator!=null&&nextBlock!=null)
      code +="not"+condition1+logicOperator;
      else 
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