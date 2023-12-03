/*
module:     src/cvcReader.ts

date:       dec-2023

usage:      utility to extract data from XLSX (excel | sheets) file
            and return the data in JSON format.
*/


import * as XLSX from 'xlsx';

export interface ExcelRow {
  [key: string]: string;
}

export class ExcelReader {
  private filePath: string;

  constructor(filePath: string) {
    this.filePath = filePath;
  }

  public async readExcelFile(): Promise<ExcelFile> {
    try {
      const workbook = XLSX.readFile(this.filePath);
      const data: ExcelRow[] = this.extractDataFromWorkbook(workbook);

      return {
        sheets: workbook.SheetNames,
        data: data,
      };
    } catch (error) {
      throw new Error(`Error reading Excel file: ${error}`);
    }
  }

  private extractDataFromWorkbook(workbook: XLSX.WorkBook): ExcelRow[] {
    const data: ExcelRow[] = [];

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const sheetData: ExcelRow[] = XLSX.utils.sheet_to_json(sheet);
      data.push(...sheetData);
    }

    return data;
  }
}

interface ExcelFile {
  sheets: string[];
  data: ExcelRow[];
}
