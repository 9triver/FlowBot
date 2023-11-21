var btn = document.getElementById("button");
    btn.onclick  =function(){  
        const code = python.pythonGenerator.workspaceToCode(workspace);
        document.getElementById('generatedCodeContainer').value = code;
        const BtnSaveAs = document.getElementById('btn-save-as')
async function fileSaveAs(description) {
    const options = {
    types: [{
        description,
        accept: {
        'text/plain': ['.txt'],
         },
     }, ],
 };
 return await window.showSaveFilePicker(options);
}
BtnSaveAs.addEventListener('click', async () => {
  const handle = await fileSaveAs("Hello File Access Api")
})
        // const blob = new Blob([code], {
        //     type: "text/plain;charset=utf-8",
        // })
        // const objectURL = URL.createObjectURL(blob);
        // objectURL.pathname="../Blockly/test";
        // const aTag = document.createElement('a');
        // aTag.href = objectURL;
        // aTag.download = "tasks.py";
        // aTag.click();
        // URL.revokeObjectURL(objectURL);
    }