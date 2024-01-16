# FlowBot
黄大伟同学负责Blockly部分，郭百川同学RPA部分
# 环境要求
Node.js 
# 安装说明
在拷贝该项目内容后，在my-electron-app目录下使用命令：
`npm install --save-dev electron`
对electron的环境要求进行安装。    
安装完毕后，在my-electron-app目录下运行：  
`npm run start`
即可开始运行该项目。  
注：electron的安装如遇其他问题，可参考官方的安装说明，以下是网址：  
https://www.electronjs.org/zh/docs/latest/tutorial/quick-start
# 额外的事项
运行和要修改的excel文件需保证在./RPA/test/data下。  
参考样例已放在sample-tasks中，里面包含样例生成的python文件和存储blockly块样例的json文件。