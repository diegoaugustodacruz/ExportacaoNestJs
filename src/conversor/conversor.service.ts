/* eslint-disable prettier/prettier */
import { Injectable } from '@nestjs/common';
import XLSXTransformStream from 'xlsx-write-stream';
const conversionFactory = require('html-to-xlsx');
const util = require('util');
const fs = require('fs');
const puppeteer = require('puppeteer');
const path = require('path');
var pdf = require('html-pdf');



@Injectable()
export class ConversorService { 

  async exportXlsx(conteudo: string) {
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

        console.log(result);

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

    const stream = await conversion(conteudo);
    stream.pipe(
      fs.createWriteStream(
        '/home/accountfy/Documentos/Buddies/htmlXlsx/teste-tabela/temp/output.xlsx',
      ),
    );
  }
  async exportPdf(){
    console.time()
    
    var html = fs.readFileSync('/home/accountfy/Documentos/Buddies/htmlXlsx/teste-tabela/temp/ex.html', 'utf8');
    var options = { format: 'Letter' };
    pdf.create(html, options).toFile('/home/accountfy/Documentos/Buddies/htmlXlsx/teste-tabela/temp/ex.pdf', function(err, res) {
      if (err) return console.log(err);
      console.log(res); // { filename: '/app/businesscard.pdf' }
    
    });
    console.timeEnd()

  }

  
}
