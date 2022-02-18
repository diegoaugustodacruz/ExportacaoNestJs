import { Body, Controller, Get, Post } from '@nestjs/common';
import { ConversorService } from './conversor.service';
const fs = require('fs');
import xlsx from 'node-xlsx';

@Controller('conversor')
export class ConversorController {
  constructor(private service: ConversorService) {}
  @Get('htmlToXlsx')
  exportXlsx() {
    this.service.exportXlsx();
    // let vetXLS = []
    // fs.readFile('/home/accountfy/projetos/ExportacaoNestJs/temp/ex.xlsx', function (err,data) {
    //   if (err) {
    //     return console.log(err);
    //   }
    //   vetXLS = data
    //   //console.log(data);
    //   console.log(typeof(data));
    // });
    var enc = new TextEncoder()
    var data = fs.readFileSync(`/home/accountfy/projetos/ExportacaoNestJs/temp/ex.xlsx`, {encoding:'utf8'})
    // Parse a buffer
    const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(`/home/accountfy/projetos/ExportacaoNestJs/temp/ex.xlsx`));
    // Parse a file
    const workSheetsFromFile = xlsx.parse(`/home/accountfy/projetos/ExportacaoNestJs/temp/ex.xlsx`);

    console.log(enc.encode(data))
    console.log(typeof(enc.encode(data)))

    return enc.encode(data);

    
  }

  @Get('htmlToPdf')
  exportPdf() {
    this.service.exportPdf();
  }

  @Get('htmlToDocx')
  exportDocx() {
    this.service.exportDocx();
  }

  @Get('htmlToPptx')
  exportPptx() {
    this.service.exportPptx();
  }
}
