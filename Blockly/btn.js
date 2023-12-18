var btn = document.getElementById("button");
var fs = require('fs');
    btn.onclick  =function(){
        var code ="from robocorp.tasks import task";
        code +="\n\nfrom ExcelExtension import ExcelApplicationExtension as ExcelApplication";
        code +="\n\n@task";
        code +="\ndef solve_challenge():\n";
        code += python.pythonGenerator.workspaceToCode(workspace);
        document.getElementById('generatedCodeContainer').value = code;
        fs.writeFile('./../RPA/test/hello.py',code);
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