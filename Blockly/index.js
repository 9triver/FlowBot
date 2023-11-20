'use strict';
let workspace = null;

function start() {
  registerFirstContextMenuOptions();
  registerExcelContent();
  workspace = Blockly.inject('blocklyDiv',
    {
        toolbox:document.getElementById('toolbox-categories'),
    });
  workspace.addChangeListener(event => {
      const code = python.pythonGenerator.workspaceToCode(workspace);
      document.getElementById('generatedCodeContainer').value = code;
    });
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
function registerExcelContent()
{ 
  var openExcel = {
    "message0": "openExcel %1",
    "args0": [
      {
      "type" : "input_value",
      "name":"FILE",
      "check":"String"
      }
    ],
    "colour":160
  };
  Blockly.Blocks['openExcel']=
    {
      init: function() {
        this.jsonInit(openExcel);
      } 
    };
    python.pythonGenerator.forBlock['openExcel'] = function(block, generator) {
      // Collect argument strings.
      const fieldValue = block.getFieldValue('MY_FIELD');
      const value = '\'' + block.getFieldValue('VALUE') + '\'';
      const innerCode = generator.statementToCode(block, 'MY_STATEMENT_INPUT');
      // Return code.
      return 'solve_challenge()';
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
