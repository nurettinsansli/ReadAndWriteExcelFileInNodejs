var xlsx = require("xlsx");
var fs = require("fs");
var path = require("path");

var sourceDir = "Files";

/*
author: nurettinsanslii@gmail.com
*/

/*
This function calculate total data count from all the excel.
*/

function readFileToJson(fileName){
    var wb = xlsx.readFile(fileName);
    var firstSheetName = wb.SheetNames[0];
    var ws = wb.Sheets[firstSheetName];
    var data = xlsx.utils.sheet_to_json(ws);
    return data;
}

var tagetDir = path.join(__dirname,sourceDir);
var files = fs.readdirSync(tagetDir);
var combinedData = [];

files.forEach(function(file){
    var fileExtenstion = path.parse(file).ext;
    if(fileExtenstion === ".xlsx" && file[0] !== "~"){
        var fullFilePath = path.join(__dirname,sourceDir,file);
        console.log(fullFilePath);
        var data = readFileToJson(fullFilePath);
        combinedData = combinedData.concat(data);
    }
});

console.log(combinedData.length);