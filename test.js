var excel = require('exceljs');
var _ = require('lodash');

var workbook = new excel.Workbook();
var sheet = workbook.addWorksheet('test');
var filename = 'test.xlsx';

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

// Column headers
sheet.columns = [
    {
        header: 'Product Id',
        key: 'pid',
        width: '25'
    },
    {
        header: 'ModCode',
        key: 'mod',
        width: '25'
    },
    {
        header: 'Pass/Fail',
        key: 'success',
        width: '25'
    }
]

// Add data programmatically
_.forEach(data, function(data) {
    sheet.addRow({
        pid: data.pid,
        mod: data.mod,
        success: data.success
    });
});

// Adding seporator for readability
sheet.addRow({pid: '------------------', mod: '------------------', success: '------------------'});

// Add data by array
sheet.addRows(data);


// Write file to disk
workbook.xlsx.writeFile(filename)
    .then(function() {
        console.log('Workbook created');
    });