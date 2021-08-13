var xlsx = require("xlsx");
var fs = require("fs");
var path = require("path");

var sourceDir = "Files";

/*
author: nurettinsanslii@gmail.com
*/

/*
This function is merge from all the excel.
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

var newWB = xlsx.utils.book_new();
var newWS = xlsx.utils.json_to_sheet(combinedData);
xlsx.utils.book_append_sheet(newWB,newWS,"Combined Data Sheet");

xlsx.writeFile(newWB,"newcombineddata.xlsx");
console.log("Done!");