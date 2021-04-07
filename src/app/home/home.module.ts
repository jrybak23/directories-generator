import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { HomeRoutingModule } from './home-routing.module';
import { HomeComponent } from './home.component';
import {ExcelReaderService} from '../excel-reader/excel-reader.service';
import {FormsModule} from '@angular/forms';


@NgModule({
  declarations: [
    HomeComponent
  ],
  imports: [
    CommonModule,
    HomeRoutingModule,
    FormsModule
  ],
  providers: [
    ExcelReaderService
  ]
})
export class HomeModule { }
