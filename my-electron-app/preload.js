const { contextBridge,ipcRenderer} = require('electron');
contextBridge.exposeInMainWorld('electronAPI', {
    node: () => process.versions.node,
    chrome: () => process.versions.chrome,
    electron: () => process.versions.electron,
    openFile: () => ipcRenderer.invoke('dialog:openFile')
    // 除函数之外，我们也可以暴露变量
  })