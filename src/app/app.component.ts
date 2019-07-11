import { Component } from '@angular/core';
import * as GC from '@grapecity/spread-sheets';
import * as Excel from '@grapecity/spread-excelio';
import '@grapecity/spread-sheets-charts';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  obj: any;
  months:any;
  spreadBackColor = 'aliceblue';
  hostStyle = {
    width: '95vw',
    height: '80vh'
  };
  private spread: GC.Spread.Sheets.Workbook;
  private excelIO;
  constructor() {
    this.excelIO = new Excel.IO();
  }
  initializeData() {
    this.obj = {
      data: [
        {
          empId: 1170906,
          empName: 'Sruthi Varghese',
          doj: '04-Sep-2018',
          cl: [0, 0, 1, 2, 3, 4, 6, 0],
          el: [12, 0, 1, 2, 1, 4, 3, 13],
          lopl: [0, 0 , 0 , 2, 0, 0, 2, -2],
        },
        {
          empId: 1170909,
          empName: 'Kuncheria Kuruvilla',
          doj: '04-Sep-2018',
          cl: [0, 1, 2, 3, 4, 4, 10, 0],
          el: [13, 1, 2, 3, 2, 4, 8, 9],
          lopl: [0, 0 , 0 , 2, 3, 0, 5, -5]
        }
      ]
    };
    this.months = this.obj.data[0].cl.length - 4;
  }
  workbookInit(args) {
    const self = this;
    self.spread = args.spread;
    const sheet = self.spread.getActiveSheet();
    sheet.getCell(0, 0).text('empID').foreColor('blue');
    sheet.getCell(1, 0).text('').foreColor('blue');
    sheet.getCell(2, 0).text('Test Excel').foreColor('blue');
    sheet.getCell(3, 0).text('Test Excel').foreColor('blue');
    sheet.getCell(0, 1).text('EmpName').foreColor('blue');
    sheet.getCell(1, 1).text('Test Excel').foreColor('blue');
    sheet.getCell(2, 1).text('Test Excel').foreColor('blue');
    sheet.getCell(3, 1).text('Test Excel').foreColor('blue');
    sheet.getCell(0, 2).text('').foreColor('blue');
    sheet.getCell(1, 2).text('Test Excel').foreColor('blue');
    sheet.getCell(2, 2).text('Test Excel').foreColor('blue');
    sheet.getCell(3, 2).text('Test Excel').foreColor('blue');
    sheet.getCell(0, 3).text('').foreColor('blue');
    sheet.getCell(1, 3).text('Test Excel').foreColor('blue');
    sheet.getCell(2, 3).text('Test Excel').foreColor('blue');
    sheet.getCell(3, 3).text('Test Excel').foreColor('blue');
 }
 onClickMe(args) {
const self = this;
  const filename = 'exportExcel.xlsx';
  const json = JSON.stringify(self.spread.toJSON());
  self.excelIO.save(json, function (blob) {
    saveAs(blob, filename);
}, function (e) {
    console.log(e);
});
}
}
