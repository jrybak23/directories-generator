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
    const beginX = this.parseFromLetter(range, 0);
    const beginY = this.parseFromNumber(range, 1);
    const endX = this.parseFromLetter(range, 3);
    const endY = this.parseFromNumber(range, 4);

    return {beginX, beginY, endX, endY};
  }

  private parseFromLetter(range: string, charIndex: number): number {
    const ALetterCharCodeOffset = 64;
    const result = range.charCodeAt(charIndex) - ALetterCharCodeOffset;
    if (!result) {
      throw new Error(`Unable to parse the letter in the range ${range}`);
    }
    return result;
  }

  private parseFromNumber(range: string, charIndex: number): number {
    const result = parseInt(range.charAt(charIndex), 10);
    if (!result) {
      throw new Error(`Unable to parse the number in the range ${range}`);
    }
    return result;
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
