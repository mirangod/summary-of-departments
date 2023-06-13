// Matheus C. @mirangod

// #region GLOBAN VAR's

const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheets = ss.getSheets();
// Status que deverão ser verificado ao decorrer do código
const status = ss.getSheetByName("VARIÁVEIS").getRange("D4:D7").getValues();
// Setores que serão verificados
const setores = ss.getSheetByName("VARIÁVEIS").getRange("F4:F17").getValues();

// #endregion

// #region SETORES

// Departamentos que passarão pela verificação
const dpto = {
  DIRECAO:    { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  COMERCIAL:  { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  COMPRAS:    { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  CUSTOS:     { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  CRIACAO:    { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  DPeRH:      { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  ENGENHARIA: { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  FINANCEIRO: { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  FISCAL:     { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  MANUTENCAO: { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  MARKETING:  { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  PCP:        { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  PRODUCAO:   { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  QUALIDADE:  { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 },
  TECNOLOGIA: { NAO_INICIADO: 0, ANDAMENTO: 0, CONCLUIDO: 0 }
}

//#endregion

/**
 * Função que percorre todas as planilhas não ocultas e verifica o departamento responsável pela atividade.
 * A coluna C é responsável por armazenar o departamento responsável.
 */
function CountSheet() { 
  for (var i = 1; i < sheets.length; i++) {
    if (!sheets[i].isSheetHidden() && sheets[i].getName() != "RESUMO" && sheets[i].getName() != "DATABASE") {
      const range = sheets[i].getRange("C:C").getValues();
      for (var j = 4; j < range.length; j++) {
        for (var k = 0; k < setores.length; k++) {
          if (range[j][0] == setores[k][0]){
            switch (setores[k][0]){
              case "DIREÇÃO":
                UploadValues(dpto.DIRECAO, sheets[i].getRange(j+1,9).getValue());
                break;
              case "COMERCIAL":
                UploadValues(dpto.COMERCIAL, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "COMPRAS":
                UploadValues(dpto.COMPRAS, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "CUSTOS":
                UploadValues(dpto.CUSTOS, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "CRIAÇÃO":
                UploadValues(dpto.CRIACAO, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "DP e RH":
                UploadValues(dpto.DPeRH, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "ENGENHARIA":
                UploadValues(dpto.ENGENHARIA, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "FINANCEIRO":
                UploadValues(dpto.FINANCEIRO, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "FISCAL":
                UploadValues(dpto.FISCAL, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "MANUTENÇÃO":
                UploadValues(dpto.MANUTENCAO, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "MARKETING":
                UploadValues(dpto.MARKETING, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "PCP":
                UploadValues(dpto.PCP, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "PRODUÇÃO":
                UploadValues(dpto.PRODUCAO, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "QUALIDADE":
                UploadValues(dpto.QUALIDADE, sheets[i].getRange(j + 1, 9).getValue());
                break;
              case "TECNOLOGIA":
                UploadValues(dpto.TECNOLOGIA, sheets[i].getRange(j + 1, 9).getValue());
                break;
              default:
                Logger.log("Setor não encontrado...");
                break;
            }
          }
        }
      }
    }
  }
  
  Logger.log(dpto);
  FillSheet();
}

/**
 * Função responsável por implementar valores no conjunto de departamentos, de acordo com o departamento encontrado na coluna C.
 * @param {object} setor - Setor que será incrementado em sua base.
 * @param {object} pesquisa - Intervalo na planilha onde será realizado a busca pelo Status afim de ser verificado.
 */
function UploadValues(setor, pesquisa) {
  switch (pesquisa) {
    case status[0][0]:
      setor.NAO_INICIADO++;
      break;
    case status[1][0]:
      setor.ANDAMENTO++;
      break;
    case status[2][0]:
      setor.CONCLUIDO++;
      break;
    default:
      Logger.log("Status não encontrado...");
      break;
  }
}

/**
 * Função que escreve os valores encontrados na função CountSheet() na planilha
 */
function FillSheet() {
  // Aba a ser depositado os valores
  const database = ss.getSheetByName("DATABASE");

  const colInicial = 0; // Coluna inicial para inserir os valores
  let row = 3; // Linha inicial para inserir os valores

  for (const dpt in dpto) {
    if (dpto.hasOwnProperty(dpt)) {
      const valores = dpto[dpt];
      const colInicialSetor = colInicial + (Object.keys(dpto[dpt]).length - 1);

      let col = colInicialSetor;
      for (const prop in valores) {
        if (valores.hasOwnProperty(prop)) {
          const valor = valores[prop];
          database.getRange(row, col).setValue(valor);
          col++;
        }
      }

      row++;
    }
  }
}

