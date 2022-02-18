import { Body, Controller, Get, Post } from '@nestjs/common';
import { readFile } from 'fs';
import { ConversorService } from './conversor.service';
const fs = require('fs');
import xlsx from 'node-xlsx';


@Controller('conversor')
export class ConversorController {
  constructor(private service: ConversorService) { }

  @Get('htmlToXlsx')
  exportXlsx() {
    this.service.exportXlsx();
  
    var enc = new TextEncoder()
    var data = fs.readFileSync(`/home/accountfy/projetos/ExportacaoNestJs/temp/ex.xlsx`, {encoding:'utf8'})
    // Parse a buffer
    const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(`/home/accountfy/projetos/ExportacaoNestJs/temp/ex.xlsx`));
    // Parse a file
    const workSheetsFromFile = xlsx.parse(`/home/accountfy/projetos/ExportacaoNestJs/temp/ex.xlsx`);

    console.log(enc.encode(data))
    console.log(typeof(enc.encode(data)))

    return enc.encode(data);

    
>>>>>>> 1ae3536c304c57873592ee1f940156d9cd3ef72d
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
