import {Component} from '@angular/core';
import {NgxDropzoneChangeEvent} from "ngx-dropzone";
import getPixels from "get-pixels";
import exceljs from 'exceljs';
import FileSaver from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  async onSelect(event: NgxDropzoneChangeEvent) {
    if (event.addedFiles.length === 0) {
      alert('Chỉ chấp nhận file jpg/jpeg/png!');
      return;
    }
    const file = event.addedFiles[0];

    const base64Image = await this.getBase64FromImage(file);
    const {width, height, imageColors} = await this.getImageDataArray(base64Image);

    let fileName = file.name;
    fileName = fileName.substring(0, fileName.lastIndexOf('.'));
    await this.drawWorkbook(width, height, imageColors, fileName);
  }

  getBase64FromImage: (file: File) => Promise<string> = (file: File) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (loadedFile) => {
        // @ts-ignore
        resolve(loadedFile.currentTarget?.result);
      };
      reader.readAsDataURL(file);
    });
  }

  getImageDataArray: (base64: string) => Promise<{ width: number, height: number, imageColors: any[] }> = (base64: string) => {
    return new Promise((resolve, reject) => {
      getPixels(base64, '', async (err, pixels) => {
        console.log('Getting image data...');
        if (err) return console.error(err);
        const {data, shape} = pixels;
        const [width, height] = shape;
        const colors = [];

        for (let i = 0; i < data.length; i += 4) {
          const a = this.numberToHexString(data[i]);
          const b = this.numberToHexString(data[i + 1]);
          const c = this.numberToHexString(data[i + 2]);
          colors.push('ff' + a + b + c);
        }
        const imageColors = [];
        while (colors.length) {
          imageColors.push(colors.splice(0, width));
        }
        resolve({width, height, imageColors});
      });
    });
  }

  drawWorkbook = async (width: number, height: number, imageColors: any[], fileName: string) => {
    console.log('Drawing...');
    const workbook = new exceljs.Workbook();
    const sheet = workbook.addWorksheet('Hi', {
      views: [
        {
          // @ts-ignore
          x: 0,
          y: 0,
          width,
          height,
          zoomScale: 50,
          showGridLines: false
        }
      ]
    });

    for (let i = 0; i < height; i++) {
      let row = [];
      for(let j = 0; j < width; j++) {
        row.push(null)
      }
      const currentRow = sheet.addRow(row);
      currentRow.height = 4;
    }

    for (let i = 0; i < width; i++) {
      const column = sheet.getColumn(i + 1);
      column.width = 1;
    }

    for (let i = 0; i < height; i++) {
      const row = sheet.getRow(i);
      for(let j = 0; j < width; j++) {
        const cell = row.getCell(j + 1);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor:{argb: imageColors[i][j]}
        }
      }
    }
    await workbook.xlsx.writeBuffer()
      .then(buffer => FileSaver.saveAs(new Blob([buffer]), `${fileName}.xlsx`));
    console.log('Done!');
  }

  numberToHexString = (number: number) => {
    const hex = number.toString(16);
    return hex.length === 1 ? `0${hex}` : hex;
  }
}
