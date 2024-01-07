'use strict';
var fs = require('fs');
fs.writeFile('./../RPA/test/hello.py', 'hello', (error) => {
  
    // 创建失败
    if(error){
        console.log(`创建失败：${error}`)
    }

    // 创建成功
    console.log(`创建成功！`)
    
})