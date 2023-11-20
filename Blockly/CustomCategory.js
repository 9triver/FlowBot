class CustomCategory extends Blockly.ToolboxCategory{
  constructor(categoryDef, toolbox, opt_parent) {
    super(categoryDef, toolbox, opt_parent);
  }
  // @override
  //   addColourBorder_(colour){
  //   this.rowDiv_.style.backgroundColor = colour;
  // }
}
Blockly.registry.register(
  Blockly.registry.Type.TOOLBOX_ITEM,
  Blockly.ToolboxCategory.registrationName,
  CustomCategory, true);