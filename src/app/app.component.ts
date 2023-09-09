import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'Oama';

  jsonData: any[] = [];
  jsonDataBlob: Blob | undefined;

  valAnaliza: number = 0;

  columnNames = [
    'Cod cerere',
    'Data cerere',
    'Nume pacient',
    'Nr.telefon',
    'Email',
    'Analize',
    'Valoare rezultat',
    'Valoare minima',
    'Valoare maxima',
    'Medic trimitator'
  ];

  onFileChange(event: any): void {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e: any) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const allSheetData: any[] = [];

      workbook.SheetNames.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        allSheetData.push({ sheetName, data: sheetData });
      });

      this.jsonData = allSheetData;

      const jsonString = JSON.stringify(this.jsonData, null, 2);
      this.jsonDataBlob = new Blob([jsonString], { type: 'application/json' });

    };
    reader.readAsArrayBuffer(file);

  }

  async blobToString(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        resolve(reader.result as string);
      };
      reader.onerror = reject;
      reader.readAsText(blob);
    });
  }

  onFormSubmit(): void {
    const containerElement = document.getElementById('table-container');

    if (!containerElement) {
      console.error(`Container element with ID "${'table-container'}" not found.`);
      return;
    }

    containerElement.innerHTML = '';


    // Iterate through each sheet's data
    this.jsonData.forEach(sheetObj => {
      const sheetName = sheetObj.sheetName;
      const sheetData = sheetObj.data;

      // Create a table element
      const table = document.createElement('table');
      table.setAttribute('border', '1'); // Add border for demonstration purposes

      // Create a table header row
      const headerRow = document.createElement('tr');

      // Iterate through the headers (assuming headers are in the first row)
      for (let i = 0; i < 10; i++) {
        const th = document.createElement('th');
        th.textContent = this.columnNames[i];
        headerRow.appendChild(th);
      }

      // Append the header row to the table
      table.appendChild(headerRow);

      // Iterate through the data rows
      for (let i = 1; i < sheetData.length; i++) {
        const rowData = sheetData[i];

        // Check if the value in the 7th cell (index 6 in a zero-based index) is greater than 4
        if (rowData.length > 6 && Number(rowData[6]) >= this.valAnaliza) {
          const row = document.createElement('tr');

          // Iterate through the data cells
          for (const cellData of rowData) {
            const cell = document.createElement('td');
            cell.textContent = cellData;
            row.appendChild(cell);
          }

          // Append the row to the table
          table.appendChild(row);
        }
      }

      // Create a heading for the table
      const tableHeading = document.createElement('h2');
      tableHeading.textContent = `Table for ${sheetName}`;

      // Append the table and heading to the container element
      containerElement.appendChild(tableHeading);
      containerElement.appendChild(table);
    });
  }


}
