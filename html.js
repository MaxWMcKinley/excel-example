var xlsx = require('xlsx');

var workbook = xlsx.readFile('exceljs.xlsx');
xlsx.writeFile(workbook, 'exceljs.html');

console.log('Html file created');
