var xlsx = require("xlsx");

function readFileToJson(fileName){
    var wb = xlsx.readFile(fileName);
    var firstSheetName = wb.SheetNames[0];
    var ws = wb.Sheets[firstSheetName];
    var data = xlsx.utils.sheet_to_json(ws);
    return data;
}

var data = readFileToJson("/Users/nurettin.sansli/Downloads/100k/1.xlsx");

console.log(data);
