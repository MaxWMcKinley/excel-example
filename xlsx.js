// This script uses xlsx to create a workbook and populate it with data
// It then saves to disk a spreadsheet in xlsx format and also an html file
var xlsx = require('xlsx');
var _ = require('lodash');

// Hardcoded example data
var data = [
    {
        pid: 'P1002650',
        mod: 'M0100',
        success: true
    },
    {
        pid: 'P1002650',
        mod: 'M0101',
        success: false
    },
    {
        pid: 'P1002650',
        mod: 'M0111',
        success: false
    },
    {
        pid: 'P1002677',
        mod: 'M0100',
        success: true
    },
    {
        pid: 'P1005600',
        mod: 'M0111',
        success: false
    }
]

// Creating worksheet
var ws = xlsx.utils.json_to_sheet(data);
var wsName = 'test';

// Create a workbook object with ws
var  wb = {
    SheetNames: [wsName],
    Sheets: {
        [wsName] : ws
    }
};

// Create a second sheet with the same data
var wsName2 = 'test2';
wb.SheetNames.push(wsName2);
wb.Sheets[wsName2] = ws;

// Write workbook to disk
xlsx.writeFile(wb, 'xlsx.xlsx');
console.log('Workbook Created');

// Write html table to disk
xlsx.writeFile(wb, 'xlsx.html');
console.log('Html File Created');