#!/usr/bin/env node
const Excel = require('exceljs');
const FormulaParser = require('hot-formula-parser').Parser;
const parser = new FormulaParser();


const excelUpdate = async (file, worksheetId, rowId, columnLabel, value) => {
   const workbook = new Excel.Workbook();
   await workbook.xlsx.readFile(file);

   const worksheet = workbook.getWorksheet(worksheetId);

   let inputRow = worksheet.getRow(rowId);
   inputRow.getCell(columnLabel).value = value;
   inputRow.commit();

   return worksheet;
}

const excelRead = async (worksheet, cellLabel) => {
   if (worksheet.getCell(cellLabel).formula) {
      return parser.parse(worksheet.getCell(cellLabel).formula).result;
    } else {
      return worksheet.getCell(cellCoord.label).value;
    }
};

const exec = async () => {
   parser.on('callCellValue', function(cellCoord, done) {
      if (worksheet.getCell(cellCoord.label).formula) {
        done(parser.parse(worksheet.getCell(cellCoord.label).formula).result);
      } else {
        done(worksheet.getCell(cellCoord.label).value);
      }
    });

   const worksheet = await excelUpdate("ExcelLogic.xlsx", "Tabelle1", 3, 'B', 5);

   const result = await excelRead(worksheet, "B5");

   console.log(result);
};

exec();