import {NgModule} from '@angular/core';
import {CommonModule} from '@angular/common';

import {HomeRoutingModule} from './home-routing.module';
import {HomeComponent} from './home.component';
import {ExcelReaderService} from '../excel-reader/excel-reader.service';
import {FormsModule} from '@angular/forms';
import {FaIconLibrary, FontAwesomeModule} from '@fortawesome/angular-fontawesome';
import { faFileDownload, faFolder } from '@fortawesome/free-solid-svg-icons';

@NgModule({
  declarations: [
    HomeComponent
  ],
  imports: [
    CommonModule,
    HomeRoutingModule,
    FormsModule,
    FontAwesomeModule
  ],
  providers: [
    ExcelReaderService
  ]
})
export class HomeModule {
  constructor(library: FaIconLibrary) {
    library.addIcons(faFolder, faFileDownload);
  }
}
