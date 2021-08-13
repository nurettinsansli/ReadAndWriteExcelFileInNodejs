var xlsx = require("xlsx");

/*
author: nurettinsanslii@gmail.com
*/

/*
This issue get data from excel. Then converted to json and write in console
*/

function readFileToJson(fileName){
    var wb = xlsx.readFile(fileName);
    var firstSheetName = wb.SheetNames[0];
    var ws = wb.Sheets[firstSheetName];
    var data = xlsx.utils.sheet_to_json(ws);
    return data;
}

var data = readFileToJson("1.xlsx");

console.log(data);
