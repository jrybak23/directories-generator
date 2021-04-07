import {Component, OnInit} from '@angular/core';
import {ExcelReaderService} from '../excel-reader/excel-reader.service';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss']
})
export class HomeComponent implements OnInit {
  private fileReader: FileReader | null = null;
  isFileChosen = false;
  items: string[] = [];
  range = '';
  message = 'No file is chosen. Choose *.xlsx file.';

  constructor(private excelReader: ExcelReaderService) {
  }

  ngOnInit(): void {
  }

  inputFile(event: Event): void {
    // @ts-ignore
    const file = event.target.files[0];
    this.fileReader = new FileReader();
    this.fileReader.onload = this.fileRead.bind(this);
    this.fileReader.readAsArrayBuffer(file);
  }

  reloadItems($event: Event): void {
    try {
      this.items = this.excelReader.readItems(1, this.range);
      if (!this.items.length) {
         this.message = `No items were found in the range ${this.range}`;
      }
    } catch (e) {
      this.items = [];
      this.message = e.message;
    }
  }

  private fileRead(progress: ProgressEvent<FileReader>): void {
    if (progress.lengthComputable) {
      this.isFileChosen = true;
      const buffer = this.fileReader?.result as ArrayBuffer;
      this.excelReader.loadWorkbook(buffer).then(() => {
        const range = this.excelReader.autoDetermineRange();
        this.range = range;
        this.items = this.excelReader.readItems(1, range);
      });
    }
  }
}
