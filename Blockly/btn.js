var btn = document.getElementById("button");
    btn.onclick  =function(){  
        alert("运行！");
        const code = python.pythonGenerator.workspaceToCode(workspace);
        eval(code);
    }