var btn = document.getElementById("button");
    btn.onclick  =function(){  
        const code = python.pythonGenerator.workspaceToCode(workspace);
        document.getElementById('generatedCodeContainer').value = code;
        const blob = new Blob([code], {
            type: "text/plain;charset=utf-8",
        })
        const objectURL = URL.createObjectURL(blob);
        objectURL.pathname="../Blockly/test";
        const aTag = document.createElement('a');
        aTag.href = objectURL;
        aTag.download = "tasks.py";
        aTag.click();
        URL.revokeObjectURL(objectURL);
    }