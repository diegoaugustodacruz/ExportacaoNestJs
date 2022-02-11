import { Body, Controller, Get, Post } from '@nestjs/common';
import { ConversorService } from './conversor.service';

@Controller('conversor')
export class ConversorController {
  constructor(private service: ConversorService) {}
  @Get('htmlToXlsx')
  exportXlsx(@Body('conteudo') conteudo: string) {
    this.service.exportXlsx(conteudo);
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
