var express = require("express");
var xlsx = require('node-xlsx');
var fs = require("fs");
var app = express();
var http = require("http");
var server = http.createServer(app);
server.listen(3000,function(){
	console.log("port 3000 has opened");
});

//向excel写入数据 
// var data = [[1,2,3,4],[true, false, null, 'sheetjs'],['foo','bar',(new Date()).toLocaleDateString()+""+(new Date()).toLocaleTimeString(), '0.3'], ['baz', null, 'qux']];  
// var buffer = xlsx.build([{name: "mySheetName", data: data}]);  
// fs.writeFileSync('b.xlsx', buffer, 'binary');  

// 读取excel中的数据
var list = xlsx.parse("b.xlsx");
console.log(list[0]);
//结果：
// { name: 'mySheetName',
//   data:
//    [ [ 1, 2, 3, 4 ],
//      [ true, false, , 'sheetjs' ],
//      [ 'foo', 'bar', '7/5/20179:25:42 PM', '0.3' ],
//      [ 'baz', , 'qux' ] ] }