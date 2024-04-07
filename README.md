
# 项目介绍
此项目是基于Blockly生成python代码内容，利用electron进行分发封装的RPA应用，用以解决繁杂重复的excel数据处理、数据清洗等问题，从而帮助NJU财务部门提高工作效率。  
Blockly是 Google 推出的一款可视化编程编辑器，使用拖放块来呈现相关的编码概念，使用所需的语言生成相对应的简洁代码，可编辑性强，灵活度高，支持生成JavaScript、python等多种不同语言的代码。  
项目的主要目标：  
1、减少相关人员的学习成本，目前主流的RPA应用门槛高、学习成本高，非编码人员使用时需要的学习周期强且操作复杂，本项目想要通过Blockly等工具降低学习门槛，将原本的命令式程序语言转化为结构化程序语言，使得没有编程背景的工作人员也能够轻松上手掌握。  
2、提高相关RPA应用的可扩展性、可维护性，开发者可自由添加相关的功能语句块，对项目进行延伸，生成自身所需的Python代码，实现所需的excel操作；  
3、加强RPA应用语句的前后关联性，项目旨在添加对不同变量的关联，减少使用者的操作量，提高效率。  
黄大伟同学负责Blockly部分，郭百川同学RPA部分
# 环境要求
Node.js Blockly
# 安装说明
在拷贝该项目内容后，在my-electron-app目录下使用命令：
`npm install --save-dev electron`
对electron的环境要求进行安装。    
安装完毕后，在my-electron-app目录下运行：  
`npm run start`
即可开始运行该项目。  
注：electron的安装如遇其他问题，可参考官方的安装说明，以下是网址：  
https://www.electronjs.org/zh/docs/latest/tutorial/quick-start
# 测试用例
测试用例由json文件构成，在应用界面选择“加载”操作即可对相同版本的json文件进行导入，生成对应的Blockly语句块。  
测试用例存放在目录sample-tasks下：  
# 额外的事项
运行和要修改的excel文件需保证在./RPA/test/data下。  
参考样例已放在sample-tasks中，里面包含样例生成的python文件和存储blockly块样例的json文件。