'use strict';

const Excel = require('exceljs');
const workbook = new Excel.Workbook();

workbook.xlsx
  .readFile('leasingberegner.xlsx')
  .then(() => {
    let js = "module.exports = function() {\nconst workbook = {};\n";
    workbook.eachSheet(sheet => {
        js += 'workbook.' + sheet.name + " = {};\n";
        sheet.eachRow({includeEmpty: true}, (row, rowNumber) => {
          row.eachCell({includeEmpty: true}, cell => {
            js += 'workbook.' + sheet.name + '.' + cell.address + ' = () => ';
            if(cell.value && cell.value.formula) {
              js += cell.value.formula.replace(/([A-Z][0-9]+)/g, 'workbook.' + sheet.name + '.$1()');
            } else {
              js += cell && cell.toCsvString() || 'null';
            }
            js += ";\n";
            for(const name of cell.names) {
              js += 'workbook.' + name + ' = ' + 'workbook.' + sheet.name + '.' + cell.address + ";\n";
            }
          });
        });
    });
    js += "return workbook;\n}\n";
    return js;
  })
  .then(sheetFormulas => process.stdout.write(sheetFormulas));
