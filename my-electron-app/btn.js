
var btnRun = document.getElementById("run");
var btnSave = document.getElementById("save");
var btnLoad =document.getElementById('load');
var btnEmpty =document.getElementById('empty');
var btnTest =document.getElementById('test');
btnTest.onclick  =function(){
    var code ="from robocorp.tasks import task";
    code +="\n\nfrom ExcelExtension import ExcelApplicationExtension as ExcelApplication";
    code +="\n\n@task";
    code +="\ndef solve_challenge():\n";
    code += python.pythonGenerator.workspaceToCode(workspace);
    document.getElementById('generatedCodeContainer').value = code;
}

btnRun.onclick  =function(){
        var code ="from robocorp.tasks import task";
        code +="\n\nfrom ExcelExtension import ExcelApplicationExtension as ExcelApplication";
        code +="\n\n@task";
        code +="\ndef solve_challenge():\n";
        code += python.pythonGenerator.workspaceToCode(workspace);
        document.getElementById('generatedCodeContainer').value = code;
        const blob = new Blob([code], {
            type: "text/plain;charset=utf-8",
        })
        
        const objectURL = URL.createObjectURL(blob);
        const aTag = document.createElement('a');
        aTag.href = objectURL;
        aTag.download = "tasks.py";
        aTag.click();
        URL.revokeObjectURL(objectURL);
    }
btnSave.onclick  =function(){
        var blockcontent=Blockly.serialization.workspaces.save(workspace);
        var myblock=JSON.stringify(blockcontent);
        const blob = new Blob([myblock], {
            type: "text/plain;charset=utf-8",
        })
        const objectURL = URL.createObjectURL(blob);
        const aTag = document.createElement('a');
        aTag.href = objectURL;
        aTag.download = "myBlock.json";
        aTag.click();
        URL.revokeObjectURL(objectURL);
        //alert("save success");
    }
btnLoad.addEventListener('click', async () => {
        const file = await window.electronAPI.openFile()
        //alert(file.toString());
        //let ob= JSON.parse(file);
        //alert(file);
        Blockly.serialization.workspaces.load(file,workspace);
        //alert(file);
})//通过监听点击事件异步加载文件内容
btnEmpty.addEventListener('click', async () => {
    const file = null;
    //alert(file.toString());
    //let ob= JSON.parse(file);
    //alert(file);
    Blockly.serialization.workspaces.load(file,workspace);
    //alert(file);
})//通
var btnLoadPath =document.getElementById('loadfilepath');
btnLoadPath.addEventListener('click', async () => {
    let path = await window.electronAPI.openFilePath();
    //alert(file.toString());
    //let ob= JSON.parse(file);
    //alert(file);
    let changePath='\'';
    var arr=new Array();
    for(let i=0;i<path.length;i++)
    {   
        
        if(path[i]=="\\")
            {   
                changePath+='/';
                //alert(path[i]);
            }
        else
            changePath+=path[i];
    }
    changePath+='\'';
    document.getElementById('generatedFilePath').value = changePath;
    //alert(file);
})
var btnLoadFolder =document.getElementById('loadfilefolder');
btnLoadFolder.addEventListener('click', async () => {
    let path = await window.electronAPI.openFileFolder();
    //alert(file.toString());
    //let ob= JSON.parse(file);
    //alert(file);
    let changePath='\'';
    var arr=new Array();
    for(let i=0;i<path.length;i++)
    {   
        
        if(path[i]=="\\")
            {   
                changePath+='/';
                //alert(path[i]);
            }
        else
            changePath+=path[i];
    }
    changePath+='/';
    changePath+='\'';
    document.getElementById('generatedFilePath').value = changePath;
    //alert(file);
})