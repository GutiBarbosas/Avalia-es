/**
 * FarmaRede — Backend de Avaliação de Desempenho
 * Google Apps Script Web App conectado a uma planilha Google Sheets.
 *
 * A planilha deve ter uma aba chamada "Avaliações" com o cabeçalho (linha 1):
 * id | funcionarioId | lojaId | mes | semana | data | gestor | scoresJson | observacao | criadoEm | atualizadoEm
 *
 * Como publicar: veja instruções no final da conversa.
 */

const NOME_ABA = 'Avaliações';
// Cole aqui o ID da planilha (está na URL dela, entre /d/ e /edit).
// Deixe assim mesmo se o script estiver vinculado à planilha (container-bound):
// usar openById remove qualquer ambiguidade sobre qual planilha é lida/gravada
// e evita o bug de "os dados somem ao atualizar a página".
const PLANILHA_ID = '17F3JJCPGgozUjnmdJwD-ATAsZV0nrDVJzJ6wRAKePMs';
const CABECALHO = [
  'id', 'funcionarioId', 'lojaId', 'mes', 'semana',
  'data', 'gestor', 'scoresJson', 'observacao', 'criadoEm', 'atualizadoEm'
];

function getPlanilha_() {
  if (PLANILHA_ID) return SpreadsheetApp.openById(PLANILHA_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'Não foi possível determinar a planilha (getActiveSpreadsheet() retornou vazio). ' +
      'Preencha a constante PLANILHA_ID no topo do Code.gs com o ID da planilha e publique uma nova versão.'
    );
  }
  return ss;
}

// Nome antigo (sem acento) que o script pode ter criado por engano em versões
// anteriores, caso NOME_ABA não batesse com a aba já existente na planilha.
const NOME_ABA_ANTIGO = 'Avaliacoes';

function getSheet_() {
  const ss = getPlanilha_();
  let sheet = ss.getSheetByName(NOME_ABA);

  // Se a aba com o nome certo não existe, procura a aba antiga (sem acento).
  // Se ela existir e tiver dados, apenas renomeia — preserva tudo que já foi salvo.
  if (!sheet) {
    const abaAntiga = ss.getSheetByName(NOME_ABA_ANTIGO);
    if (abaAntiga) {
      abaAntiga.setName(NOME_ABA);
      sheet = abaAntiga;
    }
  }

  if (!sheet) {
    sheet = ss.insertSheet(NOME_ABA);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CABECALHO);
    sheet.setFrozenRows(1);
  }
  formatarColunasComoTexto_(sheet);
  return sheet;
}

// Colunas "mes" (D) e "data" (F) guardam texto tipo "2026-07" / "2026-07-20".
// Sem isso, o Google Sheets converte esse texto automaticamente para um
// valor de Data internamente — e a comparação `avaliacao.mes === "2026-07"`
// no front-end nunca bate, fazendo o dashboard parecer "sem dados" mesmo
// com avaliações salvas.
function formatarColunasComoTexto_(sheet) {
  const COL_MES = 4;
  const COL_DATA = 6;
  const ultimaLinha = Math.max(sheet.getMaxRows(), 1000);
  sheet.getRange(2, COL_MES, ultimaLinha, 1).setNumberFormat('@');
  sheet.getRange(2, COL_DATA, ultimaLinha, 1).setNumberFormat('@');
}

// Se uma célula já foi convertida em Data (por gravações antigas antes desta
// correção), reconstrói a string "yyyy-MM" ou "yyyy-MM-dd" a partir dela.
function normalizarTexto_(valor, formato) {
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), formato);
  }
  return valor === '' || valor === null || valor === undefined ? '' : String(valor);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function linhaParaObjeto_(linha) {
  return {
    id: linha[0],
    funcionarioId: linha[1],
    lojaId: String(linha[2]),
    mes: normalizarTexto_(linha[3], 'yyyy-MM'),
    semana: Number(linha[4]),
    data: normalizarTexto_(linha[5], 'yyyy-MM-dd'),
    gestor: linha[6],
    scores: linha[7] ? JSON.parse(linha[7]) : {},
    observacao: linha[8] || '',
    criadoEm: linha[9],
    atualizadoEm: linha[10],
  };
}

/** GET — retorna todas as avaliações da planilha em JSON */
function doGet(e) {
  try {
    const sheet = getSheet_();
    const dados = sheet.getDataRange().getValues();
    const linhas = dados.slice(1).filter((l) => l[0] && l[0] !== 'id'); // ignora cabeçalho e linhas vazias/duplicadas
    const avaliacoes = linhas.map(linhaParaObjeto_);
    return jsonResponse_({ ok: true, avaliacoes, total: avaliacoes.length });
  } catch (err) {
    // Nunca deixa o erro virar a página de erro padrão do Apps Script (HTML) —
    // isso quebraria o JSON.parse no front-end e faria parecer que "os dados sumiram".
    return jsonResponse_({ ok: false, erro: String(err) });
  }
}

/**
 * Responde a requisições OPTIONS (preflight). O front-end evita preflight ao
 * enviar POST como text/plain, mas manter isso aqui é uma rede de segurança
 * caso o navegador decida mandar OPTIONS mesmo assim.
 */
function doOptions(e) {
  return jsonResponse_({ ok: true });
}

/**
 * POST — cria ou atualiza (upsert por "id") uma avaliação.
 * Corpo esperado (JSON, enviado como text/plain para evitar preflight CORS):
 * { id, funcionarioId, lojaId, mes, semana, data, gestor, scores, observacao }
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const dados = sheet.getDataRange().getValues();
    const agora = new Date().toISOString();

    let linhaExistente = -1;
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] === payload.id) { linhaExistente = i + 1; break; }
    }

    const criadoEm = linhaExistente > -1 ? dados[linhaExistente - 1][9] : agora;

    const novaLinha = [
      payload.id,
      payload.funcionarioId,
      payload.lojaId,
      payload.mes,
      payload.semana,
      payload.data,
      payload.gestor,
      JSON.stringify(payload.scores || {}),
      payload.observacao || '',
      criadoEm,
      agora,
    ];

    if (linhaExistente > -1) {
      sheet.getRange(linhaExistente, 1, 1, CABECALHO.length).setValues([novaLinha]);
    } else {
      sheet.appendRow(novaLinha);
    }

    return jsonResponse_({ ok: true, avaliacao: linhaParaObjeto_(novaLinha) });
  } catch (err) {
    return jsonResponse_({ ok: false, erro: String(err) });
  }
}

/**
 * DIAGNÓSTICO — rode esta função manualmente pelo editor do Apps Script
 * (menu "Executar" > selecionar "testarLeitura_") para confirmar, sem
 * depender do front-end nem da implantação, que o script está lendo a
 * aba certa e enxergando as avaliações salvas. Veja o resultado em
 * Ver > Registros de execução (View > Execution log).
 */
function testarLeitura_() {
  const resultado = doGet({});
  Logger.log(resultado.getContent());
}

/**
 * Igual a testarLeitura_, mas sem "_" no nome — por isso aparece no
 * dropdown de funções do editor (funções terminadas em "_" ficam ocultas
 * nesse menu). Use esta para testar pelo botão "Executar".
 */
function diagnostico() {
  const resultado = doGet({});
  Logger.log(resultado.getContent());
}

/**
 * MIGRAÇÃO — rode esta função UMA VEZ (menu "Executar" > selecionar
 * "corrigirDadosAntigos") para corrigir linhas que já foram salvas antes
 * desta correção, onde as colunas "mes" e "data" viraram Data em vez de
 * texto (ex.: "2026-07-01T03:00:00.000Z" em vez de "2026-07"). Depois de
 * rodar, confira com "diagnostico" que o campo "mes" voltou a ficar como
 * "2026-07" e o dashboard passa a exibir os dados normalmente.
 */
function corrigirDadosAntigos() {
  const sheet = getSheet_(); // já aplica o formato de texto nas colunas
  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) {
    Logger.log('Nenhuma linha de dados para corrigir.');
    return;
  }
  const COL_MES = 4;
  const COL_DATA = 6;
  const intervaloMes = sheet.getRange(2, COL_MES, ultimaLinha - 1, 1);
  const intervaloData = sheet.getRange(2, COL_DATA, ultimaLinha - 1, 1);

  const mesesCorrigidos = intervaloMes.getValues().map(
    (linha) => [normalizarTexto_(linha[0], 'yyyy-MM')]
  );
  const datasCorrigidas = intervaloData.getValues().map(
    (linha) => [normalizarTexto_(linha[0], 'yyyy-MM-dd')]
  );

  intervaloMes.setValues(mesesCorrigidos);
  intervaloData.setValues(datasCorrigidas);

  Logger.log(
    `Corrigido: ${mesesCorrigidos.length} linha(s). ` +
    `Exemplo mes: "${mesesCorrigidos[0][0]}", data: "${datasCorrigidas[0][0]}".`
  );
}
