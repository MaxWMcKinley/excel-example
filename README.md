# excel-example

This repo includes two basic examples of how to create and manipulate an excel file in node.js.  It has an example for both the xlsx and exceljs libraries.  

To set up on Unix based systems:  
1. Clone the repo
2. Run "npm install exceljs xslx lodash" to install the dependencies
3. Run either example using "node exceljs.js" or "node xlsx.js".  The worksheets should be created in the same directory.  

If you decide to make changes to the examples then close and reopen the worksheets in between uses to ensure your changes are actually written to the files.

The repo also includes a short script that will take the spreadsheet you created after running exceljs.js and output it into a html table using xlsx.  You can run this using "node html.js".  Just make sure you have run exceljs.js before using this script.
