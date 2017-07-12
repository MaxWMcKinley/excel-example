// Takes a spreadsheet and turns it into an html table
// In this case the spreadsheet is the one created by exceljs,
// so you must run that script at least once before this one
var xlsx = require('xlsx');

var workbook = xlsx.readFile('exceljs.xlsx');
xlsx.writeFile(workbook, 'exceljs.html');

console.log('Html file created');
