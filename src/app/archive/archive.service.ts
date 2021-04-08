import {Injectable} from '@angular/core';
import * as JSZip from 'jszip';
import {saveAs} from 'file-saver';

@Injectable({
  providedIn: 'root'
})
export class ArchiveService {

  constructor() {
  }

  downloadAsFoldersInZip(items: string[]): void {
    const zip = new JSZip();
    for (const item of items) {
      zip.folder(item);
    }
    zip.generateAsync({type: 'blob'})
      .then(content => {
        saveAs(content, 'result.zip');
      });
  }
}
