import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { SpreadSheetsModule } from "@grapecity/spread-sheets-angular";

import { AppComponent } from './app.component';

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule,
    SpreadSheetsModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
