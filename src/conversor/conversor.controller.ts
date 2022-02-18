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
