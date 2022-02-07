import { Injectable } from '@nestjs/common';
import XLSXTransformStream from 'xlsx-write-stream';
const conversionFactory = require('html-to-xlsx');
const util = require('util')
const fs = require('fs')
const puppeteer = require('puppeteer')
const path = require('path');


@Injectable()
export class ConversorService {
   

    async exportXlxs() {

        const writeFileAsync = util.promisify(fs.writeFile)
        const chromePageEval = require('chrome-page-eval')
        const chromeEval = chromePageEval({puppeteer});

        const conversion = conversionFactory({
            extract: async ({ html, ...restOptions }) => {
              const tmpHtmlPath = path.join('/home/accountfy/accountfy-code/teste-tabela/temp', 'input.html')
                
              console.log(tmpHtmlPath);
              writeFileAsync(tmpHtmlPath, html)
          
              const result = await chromeEval({
                ...restOptions,
                html: tmpHtmlPath,
                scriptFn: conversionFactory.getScriptFn()
              })
              
              console.log(result);
              
              const tables = Array.isArray(result) ? result : [result]
          
              return tables.map((table) => ({
                name: table.name,
                getRows: async (rowCb) => {
                  table.rows.forEach((row) => {
                    rowCb(row)
                  })
                },
                rowsCount: table.rows.length
              }))
            }
          })

          const stream = await conversion(`<table id="toExcel" class="uitable">
          <thead>
            <tr>
              <th>ECGE</th>
              <th>Janeiro</th>
              <th>Fevereiro</th>
              <th>Marco</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Salario a Pagar</td>
              <td>2000</td>
              <td>3000</td>
              <td>4000</td>
            </tr>
            <tr>
              <td>INSS</td>
              <td>50</td>
              <td>60</td>
              <td>70</td>
            </tr>
          </tbody>
        </table>`);
          stream.pipe(fs.createWriteStream('/home/accountfy/accountfy-code/teste-tabela/temp/output.xlsx'));
    }
}
