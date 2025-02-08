'use strict';
let workspace = null;
var depth=1;
var logicOperator =" null ";

const fs=require('fs');
const exec = require('child_process').exec;
const path = require('node:path');
const { app, BrowserWindow,shell,ipcMain} = require('electron');
const { dialog } = require('electron');
const { Connection } = require('blockly');
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
    item.setSavePath(process.cwd()+"\\RPA\\tasks.py")
  //item.setSavePath(__dirname+"\\RPA\\tasks.py")
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
  let myPath = "//RPA";
  let cmdStr1 = 'rcc.exe run';
  // let cmdPath = __dirname+myPath;
  let cmdPath = process.cwd()+myPath;
  // 子进程名称
  let workerProcess
  runExec(cmdStr1);
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
    workerProcess.on("close",function(code){
      console.log("out code:" + code)
      const newWin = new BrowserWindow({
        width:500,
        height:500,
    })
    newWin.loadFile(process.cwd()+"\\RPA\\output\\log.html")
    newWin.on('close',()=>{})
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
  const { canceled, filePaths } = await dialog.showOpenDialog({properties: ['openFile']})
  if (!canceled) {
    //let cur =JSON.stringify(ob);
    return filePaths[0]
}
}
async function handleFileOpenFolder () {
  const { canceled, filePaths } = await dialog.showOpenDialog({properties: ['openDirectory']})
  if (!canceled) {
    //let cur =JSON.stringify(ob);
    return filePaths[0]
}
}
app.whenReady().then(() => {
  ipcMain.handle('dialog:openFile',handleFileOpen);
  ipcMain.handle('dialog:openFilePath',handleFileOpenPath);
  ipcMain.handle('dialog:openFileFolder',handleFileOpenFolder);
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
  registerGetAllWorkbook();

  registerSetDictHeaders();
  registerMakeWorkbookDict();
  registerAddRowDict();
  registerSetDictColText();
  registerGenerateFile();

  registerMoveActiveCell();
  registerSetActiveCell();
  registerFetchCell();
  registerFetchRowNoheader();
  registerFetchRow();
  registerFetchCol();
  registerFetchAreaWithHeader();
  registerFetchArea();
  registerWriteRowNoheader();
  registerWriteRow();
  registerWriteCol();
  registerSetCellValue();

  registerCreateSheet();
  registerSetActiveSheet();
  registerMergeSheet();

  registerCompareBlock();
  registerisNoneBlock();

  registerIfBlock();
  registerElifBlock();
  registerElseBlock();
  registerSetAreaText();

  registerSettoBlock();
  registerSettoStringBlock();
  registerSettoNumBlock();

  registerGetStringLengthBlock();
  registerSubStringBlock();

  registerForBlock();
  registerForeachBlock();
  registerWhileBlock();

  registerTeleVerifyBlock();
  registerAndBlock();
  registerOrBlock();
  registerNotBlock();
  registerLeftjustBlock();

  registerInsertRowBefore();

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
    "message0": "打开工作簿：\n打开路径(%1)下的工作簿，将其命名为%2",
    "args0": [
      {
        "type": "field_input",
        "name": "FILE",
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
      FILEPATH = FILE;
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
    "message0":"保存工作簿：\n将工作簿%1保存到路径%2",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
      },
      {
        "type": "field_input",
        "name": "FILE",
        "check":"String",
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
          FILEPATH = 'filename=' + FILE ;
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
function registerGetAllWorkbook()
{ 
  var getAllWorkbook = {
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
  };
  Blockly.Blocks['getAllWorkbook']=
    {
      init: function() {
        this.jsonInit(getAllWorkbook);
      } 
    };
    python.pythonGenerator.forBlock['getAllWorkbook'] = function(block, generator) {
      // Collect argument strings.
      var VAR = block.getFieldValue('VAR');
      var Dict =block.getFieldValue('Dict');
      var FILEPATH;
      var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=VAR+"="+Dict+".names()\n";
      return code;
    }       
          
}
function registerMakeWorkbookDict()
{ 
  var MakeWorkbookDict = {
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
  };
  Blockly.Blocks['MakeWorkbookDict']=
    {
      init: function() {
        this.jsonInit(MakeWorkbookDict);
      } 
    };
    python.pythonGenerator.forBlock['MakeWorkbookDict'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
      var code='';
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
        code +=VAR+"=ExcelApplication.WorkbookDict()\n";
      return code;
    }       
          
}
function registerSetDictHeaders()
{ 
  var SetDictHeaders = {
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
  };
  Blockly.Blocks['SetDictHeaders']=
    {
      init: function() {
        this.jsonInit(SetDictHeaders);
      } 
    };
    python.pythonGenerator.forBlock['SetDictHeaders'] = function(block, generator) {
      // Collect argument strings.
      var VAR = block.getFieldValue('VAR');
      var number1= block.getFieldValue('header_row');
      if(number1!='')
      number1='header_row='+number1;
      var number2=block.getFieldValue('headers');
      if(number2!='')
      number2='headers='+number2;
      var code='';
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      if(number2=='')
      {
        if(number1=='')
        code +=VAR+".set_headers()\n";
        else
        code +=VAR+".set_headers("+number1+")\n";
      }
      else if(number1=='')
      {
        code+=VAR+".set_headers("+number2+")\n";
      }
      else
        code +=VAR+".set_headers("+number1+","+number2+")\n";
      return code;
    }       
          
}
function registerAddRowDict()
{ 
  var AddRowDict = {
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
  };
  Blockly.Blocks['AddRowDict']=
    {
      init: function() {
        this.jsonInit(AddRowDict);
      } 
    };
    python.pythonGenerator.forBlock['AddRowDict'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
      var number1= block.getFieldValue('name');
      if(number1!='')
      number1='name='+number1;
      var number2=block.getFieldValue('row_content');
      if(number2!='')
      number2='row_content='+number2;
      var code='';
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      if(number2=='')
      {
        if(number1=='')
        code +=VAR+".add_row()\n";
        else
        code +=VAR+".add_row("+number1+")\n";
      }
      else if(number1=='')
      {
        code+=VAR+".add_row("+number2+")\n";
      }
      else
        code +=VAR+".add_row("+number1+","+number2+")\n";
      return code;
    }       
          
}
function registerSetDictColText(){
  var SetDictColText ={
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
  }
  Blockly.Blocks['SetDictColText']=
    {
      init: function() {
        this.jsonInit(SetDictColText);
      } 
    };
    python.pythonGenerator.forBlock['SetDictColText'] = function(block, generator) {
      // Collect argument strings.
      const Dict = block.getFieldValue('Dict');
      var column= block.getFieldValue('column');
      column='column='+column;
      var code ="";
      for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=Dict+".column_data_type_to_text("+column+")\n";
      return code;
    }   
}
function registerGenerateFile()
{ 
  var GenerateFile = {
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
  };
  Blockly.Blocks['GenerateFile']=
    {
      init: function() {
        this.jsonInit(GenerateFile);
      } 
    };
    python.pythonGenerator.forBlock['GenerateFile'] = function(block, generator) {
      // Collect argument strings.
      const VAR=block.getFieldValue('VAR');
      var Path=block.getFieldValue('Path');
      if(Path!='')
      Path='path='+Path;
      var code='';
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code +=VAR+".generate_workbook_files("+Path+")\n";
      return code;
    }       
          
}

function registerMoveActiveCell(){
  var MoveActiveCell ={
    "message0":"移动活跃单元格：\n将工作簿%1中的活跃单元格移动%2行,%3列",
    "args0": [
      {
        "type": "field_input",
        "name": "Workbook",
        "check":"String",
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
    "tooltip":'{workbook} Move Active Cell {row_change} {column_change} : 移动活跃单元格\
    \nworkbook: Excel 文档变量名\
    \nrow_change: 行变化，默认为0\
    \ncolumn_change: 列变化，默认为0'
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
    "message0":"设置活跃单元格：\n将工作簿%1中的活跃单元格设置为第%2行,第%3列",
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
        "name":"column",
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
function registerSetAreaText(){
  var SetAreaText ={
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
  }
  Blockly.Blocks['SetAreaText']=
    {
      init: function() {
        this.jsonInit(SetAreaText);
      } 
    };
    python.pythonGenerator.forBlock['SetAreaText'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('Workbook');
      var row_from= block.getFieldValue('row_from');
      row_from='row_from='+row_from;
      var row_to= block.getFieldValue('row_to');
      row_to='row_to='+row_to;
      var column_from= block.getFieldValue('column_from');
      column_from='column_from='+column_from;
      var column_to= block.getFieldValue('column_to');
      column_to='column_to='+column_to;
      var code ="";
      for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=VAR+".data_type_to_text("+row_from+","+row_to+","+column_from+","+column_to+")\n";
      return code;
    }   
}
function registerFetchCell(){
  var fetchCell ={
    "message0":"获取单元格：\n获取工作簿%1的第%2行 第%3列，存储到变量%4中",
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
        "name":"column",
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
      var row= block.getFieldValue('row');
      if(row!='')
      row='row='+row;
      var column=block.getFieldValue('column');
      if(column!='')
      column='column='+column;
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        code +=VAR+"="+Workbook+".read_cell("+row+","+column+")\n";
      return code;
    }   
}
function registerFetchRow(){
  var fetchRow ={
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
      var header_row=block.getFieldValue('header_row');
      if(header_row!='')
      header_row='header_row='+header_row;
      var VAR=block.getFieldValue('VAR');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      // if(number4=='')
      // { 
      //   if(number3=='')
      //   {
      //     if(number2=='')
      //       {
      //       if(number1=='')
      //       code += VAR +"="+Workbook+".read_row()\n"; 
      //       else
      //       code += VAR +"="+Workbook+".read_row("+number1+")\n";
      //       }
      //     else if(number1=='')
      //     {
      //       code += VAR +"="+Workbook+".read_row("+number2+")\n";
      //     }
      //     else
      //       code += VAR +"="+Workbook+".read_row("+number1+","+number2+")\n";
      //   }
      //   else if(number2=='')
      //   { 
      //     if(number1=='')
      //       {
      //         code += VAR +"="+Workbook+".read_row("+number3+")\n";
      //       }
      //     else 
      //     code += VAR +"="+Workbook+".read_row("+number1+","+number3+")\n";
      //   }
      //   else if(number1=='')
      //   code += VAR +"="+Workbook+".read_row("+number2+","+number3+")\n";
      //   else
      //   code += VAR +"="+Workbook+".read_row("+number1+","+number2+","+number3+")\n";
      // }
      // else if(number3=='')
      // {
      //   if(number2=='')
      //   {
      //     if(number1=='')
      //     {
      //       code += VAR +"="+Workbook+".read_row("+number4+")\n";
      //     }
      //     else
      //     code += VAR +"="+Workbook+".read_row("+number1+","+number4+")\n";
      //   }
      //   else if(number1=='')
      //   {
      //     code += VAR +"="+Workbook+".read_row("+number3+")\n";
      //   }
      //   else
      //   code += VAR +"="+Workbook+".read_row("+number1+","+number2+","+number4+")\n";
      // }
      // else if(number2=='')
      // { 
      //   if(number1=='')
      //   {
      //     code += VAR +"="+Workbook+".read_row("+number3+","+number4+")\n";
      //   }
      //   else
      //   code += VAR +"="+Workbook+".read_row("+number1+","+number3+","+number4+")\n";
      // }
      // else if(number1=='')
      // { 
      //   code += VAR +"="+Workbook+".read_row("+number2+","+number3+","+number4+")\n";
      // }
      // else
      code += VAR +"="+Workbook+".read_row_with_header("+number1+","+number2+","+number3+","+header_row+")\n";
      return code;
    }   
}
function registerFetchRowNoheader(){
  var fetchRowNoheader ={
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
  }
  Blockly.Blocks['fetchRowNoheader']=
    {
      init: function() {
        this.jsonInit(fetchRowNoheader);
      } 
    };
    python.pythonGenerator.forBlock['fetchRowNoheader'] = function(block, generator) {
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
      code += VAR +"="+Workbook+".read_row("+number1+","+number2+","+number3+")\n";
      return code;
    }   
}
function registerFetchCol(){
  var fetchCol ={
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
      // if(number3=='')
      //   {
      //     if(number2=='')
      //     {
      //       if(number1=='')
      //       {
      //         code +=VAR+"="+Workbook+".read_col()\n";
      //       }
      //       else
      //       code +=VAR+"="+Workbook+".read_col("+number1+")\n";
      //     }
      //     else if(number1=='')
      //     {
      //       code +=VAR+"="+Workbook+".read_col("+number2+")\n";
      //     }
      //     else
      //     code +=VAR+"="+Workbook+".read_col("+number1+","+number2+")\n";
      //   }
      // else if(number2=='')
      // {
      //   if(number1=='')
      //   {
      //     code +=VAR+"="+Workbook+".read_col("+number3+")\n";
      //   }
      //   else 
      //   code +=VAR+"="+Workbook+".read_col("+number1+","+number3+")\n";
      // }
      // else if(number1=='')
      // {
      //   code +=VAR+"="+Workbook+".read_col("+number2+","+number3+")\n";
      // }
      // else
        code +=VAR+"="+Workbook+".read_column("+number1+","+number2+","+number3+")\n";
      return code;
    }   

}
function registerFetchAreaWithHeader(){
  var fetchArea ={
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
      var header=block.getFieldValue('header');
      header='header_row='+header;
      var VAR=block.getFieldValue('VAR');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=VAR+"="+Workbook+".read_area_wite_header("+number1+","+number2+","+number3+","+number4+","+header+")\n";
      return code;
    }   

}
function registerFetchArea(){
  var fetchArea ={
    "message0":"获取区域：\n获取工作簿%1中第%2行-%3行，第%4列-%5列的全部内容\n存储到变量%6",
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
        "check":"number",
        "name":"colF",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"colT",
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
      var VAR=block.getFieldValue('VAR');
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +=VAR+"="+Workbook+".read_area("+number1+","+number2+","+number3+","+number4+")\n";
      return code;
    }   

}
function registerWriteCol(){
  var WriteCol ={
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
  }
  Blockly.Blocks['WriteCol']=
    {
      init: function() {
        this.jsonInit(WriteCol);
      } 
    };
    python.pythonGenerator.forBlock['WriteCol'] = function(block, generator) {
      // Collect argument strings.
      const workbook = block.getFieldValue('Workbook');
      var column= block.getFieldValue('column');
      if(column!='')
      column='column='+column;
      var col_content=block.getFieldValue('col_content');
      if(col_content!='')
      col_content='column_content='+col_content;
      var row_from=block.getFieldValue('row_from');
      row_from='row_from='+row_from;
      var row_to=block.getFieldValue('row_to');
      row_to='row_to='+row_to;
      var code ="";
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      //   if(col_content=='')
      // {
      //   if(column=='')
      //   code +=workbook+".insert_column()\n";
      //   else
      //   code +=workbook+".insert_column("+column+")\n";
      // }
      // else if(column=='')
      // {
      //   code+=workbook+".insert_column("+col_content+")\n";
      // }
      // else
        code+=workbook+".write_column("+column+","+col_content+","+row_from+","+row_to+")\n";
      return code;
    }   

}
function registerInsertRowBefore(){
  var InsertRowBefore ={
    "message0":"往前插入行：\n在工作簿%1中第%2行前插入%3行新行",
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
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"num_rows",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":200,
  }
  Blockly.Blocks['InsertRowBefore']=
    {
      init: function() {
        this.jsonInit(InsertRowBefore);
      } 
    };
    python.pythonGenerator.forBlock['InsertRowBefore'] = function(block, generator) {
      // Collect argument strings.
      const workbook = block.getFieldValue('Workbook');
      var row= block.getFieldValue('row');
      if(row!='')
      row='row='+row;
      var num_rows=block.getFieldValue('num_rows');
      if(num_rows!='')
        num_rows='num_rows='+num_rows;
      var code ="";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code+=workbook+".insert_rows_before("+row+","+num_rows+")\n";
      return code;
    }   

}
function registerWriteRowNoheader(){
  var WriteRowNoheader ={
    "message0":"写入行：\n写入工作簿%1中第%2行的第%4列-%5列，行值为%3",
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
      },
      {
        "type": "field_input",
        "check":"string",
        "name":"row_content",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"column_from",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"column_to",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":'{workbook} Insert Row {row} {row_content}: 插入行\
    \nworkbook: Excel 文档变量名\
    \nrow_content: 待写入的行'
  }
  Blockly.Blocks['WriteRowNoheader']=
    {
      init: function() {
        this.jsonInit(WriteRowNoheader);
      } 
    };
    python.pythonGenerator.forBlock['WriteRowNoheader'] = function(block, generator) {
      // Collect argument strings.
      const workbook = block.getFieldValue('Workbook');
      var row= block.getFieldValue('row');
      if(row!='')
      row='row='+row;
      var row_content=block.getFieldValue('row_content');
      if(row_content!='')
      row_content='row_content='+row_content;
      var header_row=block.getFieldValue('header_row');
      if(header_row!='')
      var code ="";
      var column_from=block.getFieldValue('column_from');
      column_from='column_from='+column_from;
      var column_to=block.getFieldValue('column_to');
      column_to='column_to='+column_to;
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code+=workbook+".write_row("+row+","+row_content+","+column_from+","+column_to+")\n";
      return code;
    }   

}
function registerWriteRow(){
  var WriteRow ={
    "message0":"写入行（有表头）：\n写入工作簿%1中第%2行的第%5列-%6列，行值为%3，表头行号为%4",
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
      },
      {
        "type": "field_input",
        "check":"string",
        "name":"row_content",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"header_row",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"column_from",
      },
      {
        "type": "field_input",
        "check":"number",
        "name":"column_to",
      },
    ],
    "previousStatement": null,
    "nextStatement": null,
    "colour":160,
    "tooltip":'{workbook} Insert Row {row} {row_content} {header_row} : 插入行\
    \nworkbook: Excel 文档变量名\
    \nrow_content: 待写入的行\
    \nheader_row: header 所在行号'
  }
  Blockly.Blocks['WriteRow']=
    {
      init: function() {
        this.jsonInit(WriteRow);
      } 
    };
    python.pythonGenerator.forBlock['WriteRow'] = function(block, generator) {
      // Collect argument strings.
      const workbook = block.getFieldValue('Workbook');
      var row= block.getFieldValue('row');
      if(row!='')
      row='row='+row;
      var row_content=block.getFieldValue('row_content');
      if(row_content!='')
      row_content='row_content='+row_content;
      var header_row=block.getFieldValue('header_row');
      if(header_row!='')
      header_row='header_row='+header_row;
      var column_from=block.getFieldValue('column_from');
      column_from='column_from='+column_from;
      var column_to=block.getFieldValue('column_to');
      column_to='column_to='+column_to;
      var code ="";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      }
      code+=workbook+".write_row_with_header("+row+","+row_content+","+column_from+","+column_to+","+header_row+")\n";
      return code;
    }   

}
function registerSetCellValue(){
  var SetCellValue ={
    "message0":"设置单元格内容：\n为工作簿%1中的第%2行,第%3列，设置新值%4",
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
        "name":"column",
      },
      {
        "type": "field_input",
        "name":"value",
        "check":"string",
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
      var row= block.getFieldValue('row');
      row='row='+row;
      var column=block.getFieldValue('column');
      if(column!='')
      column='column='+column;
      var value=block.getFieldValue('value');
      value='value='+value;
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      var code ="";
      for(var i=0;i<depth;i++)
      {
        code+='    ';
      } 
      code +=VAR+".write_cell("+row+","+column+","+value+")\n";
      return code;
    }   
}
function registerSettoBlock(){
  var setto ={
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
function registerSettoStringBlock(){
  var SettoString ={
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
    "tooltip":'{workbook} Save Workbook {path} : 保存 Excel 文档 \
    \nworkbook: Excel 文档变量名 \
    \npath: 目标保存路径，为空表示在文档原位置覆盖保存'
  }
  Blockly.Blocks['SettoString']=
    {
      init: function() {
        this.jsonInit(SettoString);
      } 
    };
    python.pythonGenerator.forBlock['SettoString'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
      var exp =block.getFieldValue('exp');
        var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        code +=VAR+"="+"index_str_to_num("+exp+")\n";
      // Return code.
      return code;
    }
}
function registerSettoNumBlock(){
  var SettoNum ={
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
  }
  Blockly.Blocks['SettoNum']=
    {
      init: function() {
        this.jsonInit(SettoNum);
      } 
    };
    python.pythonGenerator.forBlock['SettoNum'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
      var exp =block.getFieldValue('exp');
        var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        code +=VAR+"="+"index_num_to_str("+exp+")\n";
      // Return code.
      return code;
    }
}
function registerGetStringLengthBlock(){
  var GetStringLength ={
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
  }
  Blockly.Blocks['GetStringLength']=
    {
      init: function() {
        this.jsonInit(GetStringLength);
      } 
    };
    python.pythonGenerator.forBlock['GetStringLength'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
      var str =block.getFieldValue('str');
        var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        code +=VAR+"="+"len("+str+")\n";
      // Return code.
      return code;
    }
}
function registerLeftjustBlock(){
  var LeftJust ={
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
  }
  Blockly.Blocks['LeftJust']=
    {
      init: function() {
        this.jsonInit(LeftJust);
      } 
    };
    python.pythonGenerator.forBlock['LeftJust'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
      var str =block.getFieldValue('str');
      var length=block.getFieldValue('Length');
        var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        code +=VAR+"=str(int("+str+")).ljust("+length+","+"'0')\n";
      // Return code.
      return code;
    }
}
function registerTeleVerifyBlock(){
  var TeleVerify ={
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
  }
  Blockly.Blocks['TeleVerify']=
    {
      init: function() {
        this.jsonInit(TeleVerify);
      } 
    };
    python.pythonGenerator.forBlock['TeleVerify'] = function(block, generator) {
      // Collect argument strings.
      const VAR = block.getFieldValue('VAR');
      var code='';
      var nextBlock =block.getNextBlock();
        code +="phone_check("+VAR+")";
      // Return code.
      return code;
    }
}
function registerSubStringBlock(){
  var SubString ={
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
  }
  Blockly.Blocks['SubString']=
    {
      init: function() {
        this.jsonInit(SubString);
      } 
    };
    python.pythonGenerator.forBlock['SubString'] = function(block, generator) {
      // Collect argument strings.
      var str = block.getFieldValue('str');
      var head =block.getFieldValue('head');
      var tail =block.getFieldValue('tail');
      var VAR = block.getFieldValue('VAR');
        var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
        code +=VAR+"="+str+"["+(head-1)+":"+tail+"]\n";
      // Return code.
      return code;
    }
}
function registerCreateSheet(){
  var CreateSheet ={
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
    "message0":"合并工作表：\n在%1工作簿里合并工作表%2和%3",
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
      {
        "type": "field_input",
        "name": "name2",
        "check":"String",
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
function registerisNoneBlock(){{
  var isNone ={
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
  }
  Blockly.Blocks['isNone']=
    {
      init: function() {
        this.jsonInit(isNone);
      } 
    };
    python.pythonGenerator.forBlock['isNone'] = function(block, generator) {
      // Collect argument strings.
      var VAR =block.getFieldValue('VAR');
      var option =block.getFieldValue('option');
      var nextBlock = block.getNextBlock();
      if(nextBlock!=null && logicOperator!=" null ")
      {
        //alert(nextBlock);
        var code =VAR+option+" None"+logicOperator;
      }
      //var prolong=generator.statementToCode(block,'prolong');
      else 
        var code =VAR+option+" None";
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
    "tooltip":' if {condition} : 分支控制块 if'
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
    "tooltip":'else if {condition} : 分支控制块 else if'
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
    'tooltip':'分支控制块 else'
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
      var End=parseInt(end);
      End=End+1;
      depth+=1;
      var content =generator.statementToCode(block,'content');
      depth-=1;
      var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +="for "+VAR+" in range"+"("+start+","+(End)+"):\n";
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
      code +="for "+VAR+" in "+iterableVar+":\n";
      code+=content;
      return code;
    }   

}}
function registerWhileBlock(){
  var While ={
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
  }
  Blockly.Blocks['While']=
    {
      init: function() {
        this.jsonInit(While);
      } 
    };
    python.pythonGenerator.forBlock['While'] = function(block, generator) {
      // Collect argument strings.
      var condition1 =generator.statementToCode(block,'condition1');
      var content =generator.statementToCode(block,'content');
      depth+=1;
      depth-=1;
      var code='';
        for(var i=0;i<depth;i++)
        {
          code+='    ';
        }
      code +="while "+condition1+":\n";
      code+=content;
      return code;
    }   

}
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
    'tooltip':'and {condition1} {condition2} ... : 与条件块'
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
    'tooltip':'or {condition1} {condition2} ... : 或条件块'
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
    'tooltip':'not {condition} : 非条件块'
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