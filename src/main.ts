/*
module:     src/main.ts

date:       dec-2023

usage:      entry-point script to test-drive util module.
*/


import { ExcelReader } from './xlsxReader.js';

const excelReader = new ExcelReader('./demo.xlsx');

excelReader.readExcelFile()
  .then((excelData) => {
    console.log('Excel Data:', excelData); })
  .catch((error) => {
    console.error('Error:', error); });
