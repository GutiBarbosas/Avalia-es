/**
 * FarmaRede — Backend de Avaliação de Desempenho
 * Google Apps Script Web App conectado a uma planilha Google Sheets.
 *
 * A planilha deve ter DUAS abas:
 *
 *   1) "base_colaboradores" — cabeçalho (linha 1):
 *      id | nome | funcao | lojaId | ativo
 *
 *   2) "Avaliações" — cabeçalho (linha 1):
 *      id | funcionarioId | lojaId | mes | semana | data | gestor | scoresJson | observacao | criadoEm | atualizadoEm
 *
 * O GET devolve { ok:true, funcionarios:[...], avaliacoes:[...] }.
 * O POST recebe uma avaliação e faz upsert (nunca duplica, nunca apaga histórico).
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

// Nome "oficial" da aba de colaboradores. Se ela não existir com este nome
// exato, tentamos localizar por nomes alternativos e, em último caso, pelo
// formato do cabeçalho (ver getAbaColaboradores_).
const NOME_ABA_COLABORADORES = 'base_colaboradores';
const CABECALHO_COLABORADORES = ['id', 'nome', 'funcao', 'lojaId', 'ativo'];

// Nomes alternativos aceitos para a aba de colaboradores, caso a planilha já
// tenha sido criada com outro nome. NUNCA inclua aqui o nome da aba de
// avaliações — isso é exatamente o bug que este arquivo corrige.
const NOMES_ALTERNATIVOS_COLABORADORES = [
  'base_colaboradores', 'Base_Colaboradores', 'Base de Colaboradores',
  'Colaboradores', 'colaboradores', 'Funcionários', 'Funcionarios',
  'funcionarios', 'funcionários', 'Equipe', 'equipe',
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
  const scores = linha[7] ? JSON.parse(linha[7]) : {};
  return {
    id: linha[0],
    funcionarioId: String(linha[1]),
    lojaId: String(linha[2]),
    mes: normalizarTexto_(linha[3], 'yyyy-MM'),
    semana: Number(linha[4]),
    data: normalizarTexto_(linha[5], 'yyyy-MM-dd'),
    gestor: linha[6],
    scores: scores,
    scoresJson: linha[7] || '',
    observacao: linha[8] || '',
    criadoEm: linha[9],
    atualizadoEm: linha[10],
  };
}

/* ============================= COLABORADORES (base_colaboradores) ============================= */

// Remove acentos e baixa a caixa, para comparar cabeçalhos de forma tolerante
// a "Loja Id", "lojaid", "LOJA_ID", etc.
function normalizarChave_(s) {
  return String(s || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

// Candidatos aceitos para cada campo do colaborador, já normalizados
// (sem acento, sem espaço, minúsculo) — cobre variações comuns de planilha.
const CAMPOS_COLABORADOR_CANDIDATOS = {
  id: ['id', 'funcionarioid', 'matricula', 'codigo', 'cod'],
  nome: ['nome', 'name', 'colaborador', 'funcionario'],
  funcao: ['funcao', 'cargo', 'role', 'funcaocargo'],
  lojaId: ['lojaid', 'loja', 'idloja', 'loja_id', 'unidade', 'unidadeid'],
  ativo: ['ativo', 'status', 'active', 'situacao'],
};

// Dado o cabeçalho (linha 1) de uma possível aba de colaboradores, tenta
// mapear cada campo esperado (id, nome, funcao, lojaId, ativo) para o índice
// da coluna correspondente. Retorna null se não achar pelo menos id e nome
// (sem isso não dá para montar um colaborador válido).
function mapearCabecalhoColaboradores_(cabecalho) {
  const normalizado = cabecalho.map(normalizarChave_);
  const mapa = {};
  Object.keys(CAMPOS_COLABORADOR_CANDIDATOS).forEach((campo) => {
    const candidatos = CAMPOS_COLABORADOR_CANDIDATOS[campo];
    let idx = -1;
    for (let i = 0; i < normalizado.length; i++) {
      if (candidatos.indexOf(normalizado[i]) > -1) { idx = i; break; }
    }
    mapa[campo] = idx;
  });
  if (mapa.id === -1 || mapa.nome === -1) return null;
  return mapa;
}

// Localiza a aba de colaboradores. Estratégia, em ordem:
//   1) nome exato NOME_ABA_COLABORADORES;
//   2) nomes alternativos conhecidos, SE o cabeçalho bater com id/nome;
//   3) varredura de todas as abas da planilha (exceto a de avaliações),
//      procurando uma cujo cabeçalho contenha pelo menos id + nome;
//   4) se nada for encontrado, cria a aba NOME_ABA_COLABORADORES vazia
//      (só com cabeçalho) — NUNCA inventa colaboradores.
// Retorna { sheet, mapa, criada } onde `mapa` é o resultado de
// mapearCabecalhoColaboradores_ (ou null se a aba acabou de ser criada).
function getAbaColaboradores_() {
  const ss = getPlanilha_();

  // 1) nome exato
  let sheet = ss.getSheetByName(NOME_ABA_COLABORADORES);
  if (sheet && sheet.getLastRow() > 0) {
    const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const mapa = mapearCabecalhoColaboradores_(cabecalho);
    if (mapa) return { sheet, mapa, criada: false };
  }

  // 2) nomes alternativos conhecidos
  for (let i = 0; i < NOMES_ALTERNATIVOS_COLABORADORES.length; i++) {
    const nome = NOMES_ALTERNATIVOS_COLABORADORES[i];
    if (nome === NOME_ABA_COLABORADORES) continue; // já testado acima
    const candidata = ss.getSheetByName(nome);
    if (candidata && candidata.getLastRow() > 0) {
      const cabecalho = candidata.getRange(1, 1, 1, candidata.getLastColumn()).getValues()[0];
      const mapa = mapearCabecalhoColaboradores_(cabecalho);
      if (mapa) return { sheet: candidata, mapa, criada: false };
    }
  }

  // 3) varredura de todas as abas, exceto a de avaliações
  const todas = ss.getSheets();
  for (let i = 0; i < todas.length; i++) {
    const candidata = todas[i];
    const nome = candidata.getName();
    if (nome === NOME_ABA || nome === NOME_ABA_ANTIGO) continue;
    if (candidata.getLastRow() === 0) continue;
    const cabecalho = candidata.getRange(1, 1, 1, candidata.getLastColumn()).getValues()[0];
    const mapa = mapearCabecalhoColaboradores_(cabecalho);
    if (mapa) return { sheet: candidata, mapa, criada: false };
  }

  // 4) nada encontrado — cria a aba oficial vazia (apenas cabeçalho).
  // Isso NUNCA apaga ou mexe em nenhuma aba existente, incluindo Avaliações.
  sheet = ss.insertSheet(NOME_ABA_COLABORADORES);
  sheet.appendRow(CABECALHO_COLABORADORES);
  sheet.setFrozenRows(1);
  return { sheet, mapa: null, criada: true };
}

function valorAtivo_(bruto) {
  const s = String(bruto === undefined || bruto === null ? '' : bruto).trim().toLowerCase();
  if (s === '') return true; // ausência de valor = considera ativo por padrão
  return !['false', '0', 'nao', 'não', 'n', 'inativo', 'no'].includes(s);
}

/**
 * Lê a aba de colaboradores e devolve um array de objetos
 * { id, nome, funcao, lojaId, ativo }. Nunca inventa colaboradores: linhas
 * sem id ou sem nome são simplesmente ignoradas. Se a aba não existir ainda
 * (foi criada agora, vazia), devolve [] — cabe a quem administra a planilha
 * preencher os dados.
 */
function getFuncionarios_() {
  const { sheet, mapa } = getAbaColaboradores_();
  if (!mapa) return []; // aba recém-criada e vazia, ou sem cabeçalho reconhecível

  const dados = sheet.getDataRange().getValues();
  const linhas = dados.slice(1); // ignora cabeçalho

  return linhas
    .map((linha) => {
      const id = linha[mapa.id];
      const nome = linha[mapa.nome];
      if (id === '' || id === null || id === undefined) return null;
      if (nome === '' || nome === null || nome === undefined) return null;
      return {
        id: String(id).trim(),
        nome: String(nome).trim(),
        funcao: mapa.funcao > -1 ? String(linha[mapa.funcao] || '').trim() : '',
        lojaId: mapa.lojaId > -1 ? String(linha[mapa.lojaId] || '').trim() : '',
        ativo: mapa.ativo > -1 ? valorAtivo_(linha[mapa.ativo]) : true,
      };
    })
    .filter(Boolean);
}

/** GET — retorna colaboradores e avaliações da planilha em JSON */
function doGet(e) {
  try {
    const sheet = getSheet_();
    const dados = sheet.getDataRange().getValues();
    const linhas = dados.slice(1).filter((l) => l[0] && l[0] !== 'id'); // ignora cabeçalho e linhas vazias/duplicadas
    const avaliacoes = linhas.map(linhaParaObjeto_);
    const funcionarios = getFuncionarios_();

    return jsonResponse_({
      ok: true,
      funcionarios,
      colaboradores: funcionarios, // alias, mantido por compatibilidade
      avaliacoes,
      total: avaliacoes.length,
    });
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
 * POST — cria ou atualiza (upsert) uma avaliação.
 * Corpo esperado (JSON, enviado como text/plain para evitar preflight CORS):
 * { id, funcionarioId, lojaId, mes, semana, data, gestor, scores, observacao }
 *
 * Chave lógica da avaliação: funcionarioId + mes + semana.
 * O front-end já envia id no formato `${funcionarioId}-${mes}-s${semana}`,
 * mas o backend NÃO confia só nisso: se não achar a linha pelo id, procura
 * também por funcionarioId+mes+semana antes de decidir criar uma linha nova.
 * Isso evita duplicar avaliações antigas que tenham sido salvas com outro
 * padrão de id.
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const dados = sheet.getDataRange().getValues();
    const agora = new Date().toISOString();

    const funcionarioId = String(payload.funcionarioId);
    const mes = String(payload.mes);
    const semana = Number(payload.semana);

    // Aceita tanto payload.scores (objeto) quanto payload.scoresJson (string ou objeto).
    let scores = payload.scores;
    if ((scores === undefined || scores === null) && payload.scoresJson) {
      scores = typeof payload.scoresJson === 'string' ? JSON.parse(payload.scoresJson) : payload.scoresJson;
    }
    scores = scores || {};

    let linhaExistente = -1;

    // 1) tenta achar pelo id enviado
    if (payload.id) {
      for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][0]) === String(payload.id)) { linhaExistente = i + 1; break; }
      }
    }

    // 2) se não achou, tenta achar pela chave lógica funcionarioId+mes+semana
    if (linhaExistente === -1) {
      for (let i = 1; i < dados.length; i++) {
        if (
          String(dados[i][1]) === funcionarioId &&
          String(dados[i][3]) === mes &&
          Number(dados[i][4]) === semana
        ) { linhaExistente = i + 1; break; }
      }
    }

    // Preserva o id e o criadoEm originais quando está atualizando uma linha
    // existente (mesmo que o front tenha enviado um id ligeiramente diferente).
    const idFinal = linhaExistente > -1
      ? dados[linhaExistente - 1][0]
      : (payload.id || `${funcionarioId}-${mes}-s${semana}`);
    const criadoEm = linhaExistente > -1 ? dados[linhaExistente - 1][9] : agora;

    const novaLinha = [
      idFinal,
      funcionarioId,
      String(payload.lojaId),
      mes,
      semana,
      payload.data,
      payload.gestor,
      JSON.stringify(scores),
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
 * depender do front-end nem da implantação, que o script está lendo as
 * abas certas e enxergando colaboradores e avaliações salvas. Veja o
 * resultado em Ver > Registros de execução (View > Execution log).
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
 * DIAGNÓSTICO — mostra especificamente qual aba foi identificada como
 * base de colaboradores e quantas linhas válidas foram lidas dela.
 * Rode pelo menu "Executar" > "diagnosticoColaboradores".
 */
function diagnosticoColaboradores() {
  const { sheet, mapa, criada } = getAbaColaboradores_();
  const funcionarios = getFuncionarios_();
  Logger.log(
    `Aba usada como base de colaboradores: "${sheet.getName()}" ` +
    `(criada agora: ${criada}). Mapeamento de colunas: ${JSON.stringify(mapa)}. ` +
    `Total de colaboradores válidos lidos: ${funcionarios.length}.`
  );
  if (funcionarios.length) {
    Logger.log(`Exemplo: ${JSON.stringify(funcionarios[0])}`);
  }
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
