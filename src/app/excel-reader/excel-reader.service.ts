import {Injectable} from '@angular/core';
import {CellValue, Row, Workbook, Worksheet} from 'exceljs';

interface ExcelRange {
  beginX: number;
  beginY: number;
  endX: number;
  endY: number;
}

@Injectable({
  providedIn: 'root'
})
export class ExcelReaderService {
  private REGEX = RegExp('([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)');
  private workbook: Workbook | null = null;

  constructor() {
  }

  loadWorkbook(buffer: ArrayBuffer): Promise<Workbook> {
    this.workbook = new Workbook();
    return this.workbook.xlsx.load(buffer);
  }

  autoDetermineRange(): string {
    const worksheet = this.workbook?.getWorksheet(1);
    const firstColumn = worksheet?.getColumn(1);
    if (!firstColumn) {
      throw Error('Unable to auto determine range. First column is not present');
    }
    const endY = firstColumn.values.length - 1;
    return `A1:A${endY}`;
  }

  readItems(worksheetIndex: number, rangeString: string): string[] {
    if (!this.workbook) {
      throw new Error('The workbook should be initialized before read.');
    }
    const range = this.parseRange(rangeString);
    const worksheet = this.workbook.getWorksheet(worksheetIndex);
    return this.readItemsFromWorksheet(worksheet, range);
  }

  private parseRange(range: string): ExcelRange {
    range = range.trim();

    const regexGroups = this.REGEX.exec(range);

    if (!regexGroups) {
      throw new Error(`Unable to parse the range ${range}. Please input range e.g A2:A11`);
    }

    // @ts-ignore
    const beginX = this.parseFromLetters(regexGroups[1]);
    // @ts-ignore
    const beginY = parseInt(regexGroups[2], 10);
    // @ts-ignore
    const endX = this.parseFromLetters(regexGroups[3]);
    // @ts-ignore
    const endY = parseInt(regexGroups[4], 10);

    return {beginX, beginY, endX, endY};
  }

  private parseFromLetters(letters: string): number {
    const column = this.workbook?.getWorksheet(1).getColumn(letters);
    // @ts-ignore
    return column?.number;
  }

  private readItemsFromWorksheet(worksheet: Worksheet, range: ExcelRange): string[] {
    const items: string[] = [];
    for (let y = range.beginY; y <= range.endY; y++) {
      for (let x = range.beginX; x <= range.endX; x++) {
        const cell = worksheet.getRow(y).getCell(x);
        if (cell.text) {
          const cellText = cell.text;
          items.push(cellText);
        }
      }
    }
    return items;
  }
}
