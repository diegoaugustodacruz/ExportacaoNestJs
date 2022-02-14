/* eslint-disable prettier/prettier */
import { Injectable } from '@nestjs/common';
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
          '/mnt/c/Users/gusta/Desktop/ExportacaoNestJs/temp',
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
        '/mnt/c/Users/gusta/Desktop/ExportacaoNestJs/temp/output.xlsx',
      ),
    );
  }

  async exportPdf(){
    console.time();

    // var html = fs.readFileSync('temp/template.html', 'utf8');
    // const html = `<html>
    //                   <head>
    //                     <title>titulo</title>
    //                     <link rel="stylesheet" href="css/style.css">
    //                     <link rel="icon" type="image/x-icon" href="images/favicon.ico">
    //                   </head>
    //                   <body>
    //                     <img src="images/1.jpg" width="100%">
    //                   </body>
    //                 </html>`

    const configNote = {
      id: 939,
      bloqueada: true,
      dataAtualizacao: "11/02/2022 14:49",
      dataCriacao: "09/02/2022 15:24",
      dataReferencia: "01/03/2022",
      idConteudoNota: 455,
      idNotaCabecalhoRodape: 455,
      grupoEconomicoId: 1056,
      tipoApresentacaoNota: "DOCUMENTO",
      idUsuario: 3830,
      idUsuarioAlteracao: 4908,
      titulo: "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
      fonte: 1,
      tamanhoFonte: 1,
      margin: 1,
      sumario: false,
      listaPrimaria: 1,
      listaSecundaria: 2
  }

    const cabecalhoERodape = {
      id:455,
      conteudo: {
        noteFooter: true,
        noteHeader: true,
        footerConfig: [
          {
            col:"4",
            type:"pagination",
            align:"left",
            hidden:false
          }, {
            col:"5",
            type:"text",
            align:"center",
            hidden:false
          }, {
            col:"6",
            type:"text",
            align:"right",
            hidden:false
          }],
          headerConfig: [
          {
            col:"1",
            type:"text",
            align:"left",
            hidden:false
          }, {
            col:"2",
            type:"text",
            align:"center",
            hidden:false
          }, {
            col:"3",
            type:"text",
            align:"right",
            hidden:false
          }],
          footerPagination: false,
          numberColumnsFooter: 3,
          numberColumnsHeader: 3
        }
    }

    const templateCabecalho = '';
    

    // header
    let htmlHeader: any;
    let colsHeader;
    cabecalhoERodape.conteudo.headerConfig.forEach((col) => {
      if (!col.hidden) {
        colsHeader += `<div class="flex-grow-1 item-padding table-styles">
                        <div class="area-editable item-block position-relative d-flex page-var table-styles w-100 h-100">
                          <div class="item w-100 h-100">
                            ${col.col}
                          </div>
                        </div>
                      </div>`;
      }
    });
    htmlHeader = cabecalhoERodape.conteudo.noteHeader ? {
      height: "80px",
      contents: `<div class="item-dragula item-dragula-row position-relative page-cabecalho">
      <div class="pageWrapper-item position-relative">
        <div class="pw-item d-flex align-items-stretch item-dragula-col">
          ${colsHeader}
        </div>
      </div>
    </div>`
    } : false;

    // rodape
    let htmlFooter: any;
    let colsFooter;

    cabecalhoERodape.conteudo.footerConfig.forEach((col) => {
      if (!col.hidden) {
        colsFooter += `<div class="item">
                        ${col.type === 'pagination' && cabecalhoERodape.conteudo.footerPagination ? '' : col.col}
                      </div>`;
      }
    });
    htmlFooter = cabecalhoERodape.conteudo.noteHeader ? {
      height: "40px",
      contents: `<div class="page-rodape pageWrapper-item">${colsFooter}</div>`,
    } : false;

    const objt = {
      tipoDownload: 1,
      pages: [
          {
              id: 1,
              size: 1,
              type: 1,
              layout: 2,
              rows: [
                  {
                      align: "left",
                      cols: [
                          {
                              htmlContent: "1",
                              align: "left"
                          }
                      ]
                  },
                  {
                      align: "left",
                      cols: [
                          {
                              htmlContent: "1",
                              align: "left"
                          },
                          {
                              htmlContent: "2",
                              align: "left"
                          }
                      ]
                  },
                  {
                      align: "left",
                      cols: [
                          {
                              htmlContent: "1",
                              align: "left"
                          },
                          {
                              htmlContent: "2",
                              align: "left"
                          },
                          {
                              htmlContent: "3",
                              align: "left"
                          }
                      ]
                  },
                  {
                      align: "left",
                      cols: [
                          {
                              exceptionMeasure: 1.5,
                              htmlContent: "12",
                              align: "left"
                          },
                          {
                              exceptionMeasure: 3,
                              htmlContent: "3",
                              align: "left"
                          }
                      ]
                  },
                  {
                      align: "left",
                      cols: [
                          {
                              exceptionMeasure: 3,
                              htmlContent: "1",
                              align: "left"
                          },
                          {
                              exceptionMeasure: 1.5,
                              htmlContent: "23",
                              align: "left"
                          }
                      ]
                  },
                  {
                      id: null,
                      noteId: null,
                      note: true,
                      cols: null,
                      titulo: "nota",
                      ordem: 1,
                      align: "left",
                      descricao: [
                          {
                              htmlContent: "nota",
                              align: "left"
                          }
                      ],
                      estruturas: [],
                      notasSecundarias: [
                          {
                              align: "left",
                              id: null,
                              noteId: null,
                              note: true,
                              cols: null,
                              titulo: "nota sec",
                              ordem: 1,
                              descricao: [
                                  {
                                      htmlContent: "nota sec",
                                      align: "left"
                                  }
                              ],
                              estruturas: [],
                              notasSecundarias: []
                          }
                      ]
                  },
                  {
                      align: "left",
                      cols: [
                          {
                              htmlContent: "\n      <div contenteditable=\"false\" class=\"tabela-notas position-relative \" data-consider-dtr=\"true\" data-tabela-ref=\"1781\" id=\"tabela-ref-1781\">\n        <div class=\"tabela-notas-config flex\">\n          <span>Balanço Patrimonial (versão Notas Explicativas)</span>\n          <button class=\"btn btn-note btn-refresh-table ml-10\" data-table-id=\"1781\"><span class=\"icon-refresh\"></span></button>\n          <button class=\"btn btn-note btn-delete-table\" data-table-id=\"1781\"><span class=\"icon-delete\"></span></button>\n          <button class=\"btn btn-note btn-edit-table\" data-table-id=\"1781\"><span class=\"icon-edit\"></span></button>\n        </div>\n        <div class=\"table-item table-item-multiple table-item-multiple-left\">\n              <table class=\"table table-ativo-passivo\">\n                <thead class=\"label-header\">\n                  <tr class=\"underline\">\n                    <td></td>\n                    <td></td>\n                    \n                    \n                  </tr>\n\n                  <tr class=\"line-father-total\">\n                    <td></td>\n                    <td></td>\n                    \n                    <td>31/03/2022</td>\n                    <td></td>\n                    <td>31/03/2021</td>\n                  </tr>\n                  <tr>\n                    <td><strong>ATIVO</strong></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr>\n\n                </thead><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Caixa e Equivalentes\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">1,101,233</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">1,101,233</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Títulos e valores mobiliários\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">92,029</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">92,029</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Garantia de Investimentos\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">9,718</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">9,718</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Contas a Receber\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">1,929,661</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">1,929,661</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Provisão para Crédito de Liquidação Duvidosa\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(141,350)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(141,350)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Tributos a recuperar\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">37,005</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">37,005</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Outros Ativos CP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">63,835</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">63,835</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"line-father-total\">\n                <td class=\"\">\n                  Ativo Circulante\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">3,092,132</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">3,092,132</td>\n              </tr>\n              \n                <tr>\n                  <td></td>\n                  <td></td>\n                  \n                  <td></td>\n                  <td></td>\n                  <td></td>\n                </tr>\n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Contas a Receber de Clientes LP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">55,988</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">55,988</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Imposto de Renda e Contribuição Social Diferidos LP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">115,630</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">115,630</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Depósitos Judiciais\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">45,335</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">45,335</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Ativo Financeiro a Valor Justo\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">101,706</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">101,706</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Outros Ativos LP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">65,090</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">65,090</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Investimentos\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">3,707</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">3,707</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Imobilizado\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">394,934</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">394,934</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Intangível\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">1,562,809</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">1,562,809</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"line-father-total\">\n                <td class=\"\">\n                  Ativo Não Circulante\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">2,345,200</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">2,345,200</td>\n              </tr>\n              \n                <tr>\n                  <td></td>\n                  <td></td>\n                  \n                  <td></td>\n                  <td></td>\n                  <td></td>\n                </tr>\n            </tbody><tfoot class=\"label-footer\"><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr><tr>\n                    <td></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr>\n        <tr>\n          <td></td>\n          <td></td>\n          \n          <td></td>\n          <td></td>\n          <td></td>\n        </tr>\n        <tr>\n          <td>Total do Ativo</td>\n          <td></td>\n          \n          <td class=\"celula-borda-b celula-borda-t\" data-testid=\"valorReferenciaConsolidado\">\n            5,437,332\n          </td>\n          <td class=\"celula-borda-b celula-borda-t\"></td>\n          <td class=\"celula-borda-b celula-borda-t\" data-testid=\"valorComparacaoConsolidado\">\n            5,437,332\n          </td>\n        </tr></tfoot></table></div><div class=\"table-item table-item-multiple table-item-multiple-right\">\n              <table class=\"table table-ativo-passivo\">\n                <thead class=\"label-header\">\n                  <tr class=\"underline\">\n                    <td></td>\n                    <td></td>\n                    \n                    \n                  </tr>\n\n                  <tr class=\"line-father-total\">\n                    <td></td>\n                    <td></td>\n                    \n                    <td>31/03/2022</td>\n                    <td></td>\n                    <td>31/03/2021</td>\n                  </tr>\n                  <tr>\n                    <td><strong>PASSIVO E PL</strong></td>\n                    <td></td>\n                    \n                    <td></td>\n                    <td></td>\n                    <td></td>\n                  </tr>\n\n                </thead><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Obrigações Sociais e Trabalhistas\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(215,679)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(215,679)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Fornecedores\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(77,800)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(77,800)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Obrigações Fiscais\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(70,708)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(70,708)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Empréstimos e Financiamentos CP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(156,139)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(156,139)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Comissões a Pagar\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(61,738)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(61,738)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Dividendos e JCP a Pagar\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(51,659)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(51,659)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Obrigações por Aquisição de Investimentos CP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(29,400)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(29,400)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Repasse para Parceiros\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(381,622)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(381,622)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Cota Seniores e Mezanino\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(1,135,225)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(1,135,225)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Outros Passivos CP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(11,802)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(11,802)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"line-father-total\">\n                <td class=\"\">\n                  Passivo Circulante\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(2,191,772)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(2,191,772)</td>\n              </tr>\n              \n                <tr>\n                  <td></td>\n                  <td></td>\n                  \n                  <td></td>\n                  <td></td>\n                  <td></td>\n                </tr>\n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Empréstimos e Financiamentos LP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(205,476)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(205,476)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Prov. para Obrig. Legais Vinc. a Processos Judiciais\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(135,818)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(135,818)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Obrigações por Aquisição de Investimentos LP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(157,653)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(157,653)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Obrigações Fiscais LP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(3,672)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(3,672)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Imposto de Renda e Contribuição Social Diferidos LP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(2,004)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(2,004)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Outros Passivos LP\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(37,418)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(37,418)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"line-father-total\">\n                <td class=\"\">\n                  Exigível a Longo Prazo\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(542,041)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(542,041)</td>\n              </tr>\n              \n                <tr>\n                  <td></td>\n                  <td></td>\n                  \n                  <td></td>\n                  <td></td>\n                  <td></td>\n                </tr>\n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Capital Social\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(1,382,824)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(1,382,824)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Ações em Tesouraria\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">148,478</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">148,478</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Reserva de Capital\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(901,880)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(901,880)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Reserva de Lucros\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(381,849)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(381,849)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Outros resultados abrangentes\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(53,303)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(53,303)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Proposta de Dividendos Adicionais\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(50,960)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(50,960)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"\">\n                <td class=\"\">\n                  Resultados do Exercício\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(81,181)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(81,181)</td>\n              </tr>\n              \n            </tbody><tbody>\n              <tr class=\"line-father-total\">\n                <td class=\"\">\n                  Patrimônio Líquido\n                </td>\n                <td></td>\n                \n                <td data-testid=\"valorReferenciaConsolidado\">(2,703,519)</td>\n                <td></td>\n                <td data-testid=\"valorComparacaoConsolidado\">(2,703,519)</td>\n              </tr>\n              \n                <tr>\n                  <td></td>\n                  <td></td>\n                  \n                  <td></td>\n                  <td></td>\n                  <td></td>\n                </tr>\n            </tbody><tfoot class=\"label-footer\">\n        <tr>\n          <td></td>\n          <td></td>\n          \n          <td></td>\n          <td></td>\n          <td></td>\n        </tr>\n        <tr>\n          <td>Total do Passivo e PL</td>\n          <td></td>\n          \n          <td class=\"celula-borda-b celula-borda-t\" data-testid=\"valorReferenciaConsolidado\">\n            (5,437,332)\n          </td>\n          <td class=\"celula-borda-b celula-borda-t\"></td>\n          <td class=\"celula-borda-b celula-borda-t\" data-testid=\"valorComparacaoConsolidado\">\n            (5,437,332)\n          </td>\n        </tr></tfoot></table></div>\n      </div>",
                              align: "left"
                          }
                      ]
                  }
              ]
          }
      ]
    }

    const fontSizeClass = configNote.tamanhoFonte === 2 ? 'font-medium' : configNote.tamanhoFonte === 3 ? 'font-large' : 'font-small';
    let templatePage = ``;

    objt.pages.forEach(page => {
      let templateRow = '';

      page.rows.forEach(row => {
        let templateCol = '';

       if (!row.note) {
        row.cols.forEach((col) => {
          templateCol += `<div class="flex-grow-1 item-padding table-styles" ${col.exceptionMeasure && 'style="min-width: calc(100% / ' + col.exceptionMeasure + ');"'}>
                            <div class="item-block position-relative d-flex page-var table-styles">
                              <div class="item w-100 font-${configNote.fonte} ${fontSizeClass}">
                                ${col.htmlContent}
                              </div>
                            </div>
                          </div>`;
        });

        templateRow += `<div class="item-dragula item-dragula-row position-relative table-styles">
                          <div class="pageWrapper-item position-relative">
                            <div class="pw-item d-flex align-items-stretch item-dragula-col">
                              ${templateCol}
                            </div>
                          </div>
                        </div>`;
        }
      });

      let fontFamily = '';
      let fontFace = '';
      if (configNote.fonte === 1) {
        fontFamily = 'Open-Sans'
        fontFace = `@font-face {
                      font-family: "Open-Sans";
                      font-weight: 400;
                      src: url("fonts/OpenSans/OpenSans-Regular.ttf");
                    }`
      } else if (configNote.fonte === 2) {
        fontFamily = 'Roboto'
        fontFace = `@font-face {
                      font-family: "Roboto";
                      font-weight: 400;
                      src: url("fonts/Roboto/Roboto-Regular.ttf");
                    }`
      } else if (configNote.fonte === 3) {
        fontFamily = 'Arimo'
        fontFace = `@font-face {
                      font-family: "Arimo";
                      font-weight: 400;
                      src: url("fonts/Arimo/Arimo-Regular.ttf");
                    }`
      } else if (configNote.fonte === 4) {
        fontFamily = 'Fira-Sans'
        fontFace = `@font-face {
                      font-family: "Fira-Sans";
                      font-weight: 400;
                      src: url("fonts/FiraSans/FiraSans-Regular.ttf");
                    }`
      }  else if (configNote.fonte === 5) {
        fontFamily = 'Garamond'
        fontFace = `@font-face {
                      font-family: "Garamond";
                      font-weight: 400;
                      src: url("fonts/Garamond/Garamond-Regular.ttf");
                    }`
      }

      templatePage = `
          <!DOCTYPE html>
          <html lang="en">
            <head>
              <meta charset="utf-8">
              <meta name="viewport" content="width=device-width, initial-scale=1" />
              <title>${configNote.titulo}</title>
              <link rel="stylesheet" href="css/style.css">

              <style>
                ${fontFace}
              *,html,body {
                font-family: ${fontFamily} !important;
              }
              </style>
            </head>
            <body class="area-content note-container table-styles font-${configNote.fonte} ${fontSizeClass}">
              <section class="item-content position-relative d-flex align-items-start">
                <div class="pageWrapper pageWrapper-valid page-content">
                  <div class="pageWrapper-block">
                    <div class="container-dragula container-dragula-row position-relative">
                      ${templateRow}
                    </div>
                  </div>
                </div>
              </section>
            </body>
          </html>
        `;

        const base = path.resolve('static');
    
        const format = page.size === 1 ? 'A4' : 'A3';
        const orientation = page.layout === 1 ? 'portrait' : 'landscape';
        const margin = configNote.margin === 1 ? '30px' : configNote.margin === 2 ? '50px' : '70px';
    
        const options = { 
          format: format, // 1 = A3, 2 = A4
          orientation: orientation, // 1 = portrait  2 = landscape
          border: margin, // 30, 50, 70
          // paginationOffset: 1,
          header: htmlHeader,
          footer: htmlFooter,
          base: `file://${base}/`,
          // phantomPath: "./node_modules/phantomjs-prebuilt/bin/phantomjs", // PhantomJS binary which should get downloaded automatically
          localUrlAccess: true, // Prevent local file:// access by passing '--local-url-access=false' to phantomjs
          // timeout: 30000,        // Timeout that will cancel phantomjs, in milliseconds
        };
    
        const nome = new Date();
        pdf.create(templatePage, options).toFile(`temp/${nome}.pdf`, function(err, res) {
          console.log(templatePage);
          
          if (err) return console.log(err);
          console.log(res);
        });
    
        console.timeEnd();
    });

  }
}
