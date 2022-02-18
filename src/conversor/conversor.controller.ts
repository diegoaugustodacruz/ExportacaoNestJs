import { Body, Controller, Get, Post } from '@nestjs/common';
import { readFile } from 'fs';
import { ConversorService } from './conversor.service';
const fs = require('fs');
import xlsx from 'node-xlsx';

@Controller('conversor')
export class ConversorController {
  constructor(private service: ConversorService) {}

  @Get('htmlToXlsx')
  exportXlsx() {
    this.service.exportXlsx();

    const enc = new TextEncoder();
    const data = fs.readFileSync(
      `/home/accountfy/Documentos/Buddies/htmlXlsx/teste-tabela/temp/ex.xlsx`,
      { encoding: 'utf8' },
    );

    const emBas64 = new Buffer(data).toString('base64');

    console.log(emBas64);

    return emBas64;
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
