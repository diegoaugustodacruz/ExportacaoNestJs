/* eslint-disable prettier/prettier */
import { Injectable } from '@nestjs/common';
import XLSXTransformStream from 'xlsx-write-stream';
const conversionFactory = require('html-to-xlsx');
const util = require('util');
const fs = require('fs');
const puppeteer = require('puppeteer');
const path = require('path');
var pdf = require('html-pdf');
const HTMLtoDOCX = require('html-to-docx');
const html2pptxgenjs = require('html2pptxgenjs');
const pptxgen = require('pptxgenjs');

@Injectable()
export class ConversorService {
  async exportXlsx() {
    console.time();
    const chromePageEval = require('chrome-page-eval');
    const writeFileAsync = util.promisify(fs.writeFile);
    const chromeEval = chromePageEval({ puppeteer });
    

    const conversion = conversionFactory({
      extract: async ({ html, ...restOptions }) => {
        const tmpHtmlPath = path.join(
          '/home/accountfy/Documentos/Buddies/htmlXlsx/teste-tabela/temp',
          'input.html',
        );

        console.log(tmpHtmlPath);
        writeFileAsync(tmpHtmlPath, html);

        const result = await chromeEval({
          ...restOptions,
          html: tmpHtmlPath,
          scriptFn: conversionFactory.getScriptFn(),
        });

        const tables = Array.isArray(result) ? result : [result];

        return tables.map((table) => ({
          name: table.name,
          getRows: async (rowCb) => {
            table.rows.forEach((row) => {
              rowCb(row);
            });
          },
          rowsCount: table.rows.length,
        }));
      },
    });
        
    const html = fs.readFileSync('temp/ex.html', 'utf8');

    const stream = await conversion(html);
    stream.pipe(
      fs.createWriteStream(
        '/home/accountfy/Documentos/Buddies/htmlXlsx/teste-tabela/temp/ex.xlsx',
      ),
    );
    console.timeEnd();

  }
  async exportPdf() {
    console.time();

    var html = fs.readFileSync('temp/ex.html', 'utf8');
    var options = {
      format: 'Letter',
      height: ' 297mm', // allowed units: mm, cm, in, px
      width: '210mm', // allowed units: mm, cm, in, px
      orientation: 'portrait', // portrait or landscape
      border: {
        top: '15px', // default is 0, units: mm, cm, in, px
        right: '15px',
        bottom: '15px',
        left: '15px',
      },
    };
    pdf.create(html, options).toFile('temp/ex.pdf', function (err, res) {
      if (err) return console.log(err);
      console.log(res); // { filename: '/app/businesscard.pdf' }
    });
    console.timeEnd();
  }

  async exportDocx() {
    console.time();

    const filePath =
      '/home/accountfy/Documentos/Buddies/htmlXlsx/teste-tabela/temp/ex.docx';

    const htmlString = fs.readFileSync('temp/ex.html', 'utf8');

    (async () => {
      const fileBuffer = await HTMLtoDOCX(htmlString, null, {
        table: { row: { cantSplit: true } },
        header: false,
        footer: false,
        pageNumber: false,
      });

      fs.writeFile(filePath, fileBuffer, (error) => {
        if (error) {
          console.log('Docx file creation failed');
          return;
        }
        console.log('Docx file created successfully');
      });
    })();

    console.timeEnd();
  }

  async exportPptx() {
    const html = fs.readFileSync('temp/demo2.html', 'utf8');
    const Slides = [html];
    const options = {
      css: `
      * {
        text-decoration: none !important;
      }`
    };

    
    const pres = new pptxgen();

    Slides.forEach(text => {
      const slide = pres.addSlide();

      const items = html2pptxgenjs.htmlToPptxText(text, options);

      console.log(items)
      
      slide.addText(items, { x: 0.5, y: 0, w: 9.5, h: 6, valign: 'top' });
      
      return items;
    });

    pres.writeFile('temp/demo2.pptx');

  }
}