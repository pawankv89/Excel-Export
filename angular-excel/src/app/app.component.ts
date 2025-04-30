import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  title = 'angular-excel';

  exportExcelFile() {
    const workbook = new ExcelJS.Workbook();
    const sheetData = this.getsheetData();
    sheetData.forEach((sheet: any) => {
      const worksheet = workbook.addWorksheet(sheet?.sheetName);
      if (sheet?.data?.length > 0) {
        const header = Object.keys(sheet?.data[0]);
        worksheet.addRow(header);
        sheet?.data?.forEach((rowData: any) => {
          const row = Object.values(rowData);
          worksheet.addRow(row);
        });
      }
    });
    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, 'Student.xlsx');
    });
  }

  getsheetData() {
    const items: any = [
      {
        "sheetName": "Pawan",
        "data" : [
          {
            "Name": "Pawan",
            "Age": "30",
            "Address": "Sector 50",
            "Location": "Noida"
          }
        ]
      },
      {
        "sheetName": "Kumar",
        "data" : [
          {
            "Name": "Kumar",
            "Age": "29",
            "Address": "Sector 49",
            "Location": "Noida"
          }
        ]
      },
      {
        "sheetName": "Sharma",
        "data" : [
          {
            "Name": "Sharma",
            "Age": "28",
            "Address": "Sector 48",
            "Location": "Noida"
          }
        ]
      }
    ]
    return items;
  }
}
