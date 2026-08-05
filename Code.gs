/**
 * ============================================================
 *  SISTEMA DE REABERTURAS — Cemitério da Saudade
 *  Backend em Google Apps Script
 * ============================================================
 *
 *  O que este código faz:
 *  Ele transforma esta planilha em um "servidor" que o sistema
 *  (a página web) usa para salvar e ler os dados. Assim, quando
 *  qualquer pessoa da equipe abre o sistema em qualquer
 *  computador, todos veem e alteram os MESMOS dados.
 *
 *  Este código cria duas abas novas na planilha, separadas das
 *  suas abas antigas (REABERTURA 2022, 2023...), chamadas:
 *    - BASE_SISTEMA   -> onde ficam todos os registros ativos
 *    - LOG_SISTEMA    -> um histórico simples de atividades
 *
 *  Suas abas antigas (REABERTURA 2022 a 2029) NÃO são apagadas
 *  nem alteradas. Elas continuam aí como estavam.
 *
 *  PASSO OBRIGATÓRIO ANTES DE USAR:
 *  Troque a senha abaixo (linha "const TOKEN") por uma senha
 *  só sua. Essa mesma senha você vai colar depois dentro do
 *  arquivo HTML do sistema. Ela impede que qualquer pessoa na
 *  internet consiga ler ou apagar os dados do cemitério.
 * ============================================================
 */

// ⚠️ TROQUE ESTA SENHA por uma só sua (letras e números, sem espaço)
const TOKEN = 'TROQUE-ESTA-SENHA-2026';

const SHEET_BASE = 'BASE_SISTEMA';
const SHEET_LOG  = 'LOG_SISTEMA';


/* ============================================================
 *  PONTOS DE ENTRADA
 *  O sistema web "conversa" com estas duas funções.
 * ============================================================ */

function doGet(e) {
  return tratarPedido(e);
}

function doPost(e) {
  return tratarPedido(e);
}

function tratarPedido(e) {
  try {
    const parametros = (e && e.parameter) || {};

    let corpo = {};
    if (e && e.postData && e.postData.contents) {
      try { corpo = JSON.parse(e.postData.contents); } catch (erro) { corpo = {}; }
    }

    const senhaRecebida = parametros.token || corpo.token;
    if (senhaRecebida !== TOKEN) {
      return responderJSON({ ok: false, error: 'Acesso negado. Senha incorreta ou ausente.' });
    }

    const acao = parametros.action || corpo.action;

    if (acao === 'ping') {
      return responderJSON({ ok: true, agora: new Date().toISOString(), mensagem: 'Conectado com sucesso.' });
    }

    if (acao === 'get') {
      const chave = parametros.key || corpo.key;
      return responderJSON(lerChave(chave));
    }

    if (acao === 'set') {
      const chave = corpo.key;
      const valor = corpo.value;
      return responderJSON(gravarChave(chave, valor));
    }

    // Grava (ou atualiza) UM ÚNICO registro, sem mexer nos demais.
    // É isso que permite a equipe inteira editar casos diferentes ao
    // mesmo tempo sem um salvamento apagar o do outro.
    if (acao === 'upsert') {
      return responderJSON(gravarUmRegistro(corpo.record));
    }

    // Remove um único registro pelo id.
    if (acao === 'remove') {
      return responderJSON(removerUmRegistro(corpo.id));
    }

    // Acrescenta uma linha ao histórico de atividades sem apagar as
    // anteriores.
    if (acao === 'appendLog') {
      return responderJSON(acrescentarLog(corpo.text));
    }

    return responderJSON({ ok: false, error: 'Ação não reconhecida: ' + acao });

  } catch (erro) {
    return responderJSON({ ok: false, error: 'Erro no servidor: ' + String(erro) });
  }
}

function responderJSON(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ============================================================
 *  ARMAZENAMENTO
 *  Cada registro (cada reabertura) vira uma linha na aba
 *  BASE_SISTEMA, guardando o registro inteiro como texto (JSON)
 *  na coluna B. Isso evita problemas de formatação e permite
 *  qualquer campo novo no futuro sem precisar mexer na planilha.
 * ============================================================ */

function nomeDaAbaParaChave(chave) {
  if (chave === 'reaberturas:records') return SHEET_BASE;
  if (chave === 'reaberturas:log')     return SHEET_LOG;
  return null;
}

function obterOuCriarAba(nome) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(nome);
  if (!aba) {
    aba = planilha.insertSheet(nome);
    aba.appendRow(['id', 'json']);
    aba.setFrozenRows(1);
    aba.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  return aba;
}

function lerChave(chave) {
  const nomeAba = nomeDaAbaParaChave(chave);
  if (!nomeAba) return { ok: false, error: 'Chave desconhecida: ' + chave };

  const aba = obterOuCriarAba(nomeAba);
  const dados = aba.getDataRange().getValues();
  const linhas = dados.slice(1).filter(function (linha) { return linha[1]; });

  if (linhas.length === 0) {
    return { ok: true, found: false };
  }

  let valor;
  if (chave === 'reaberturas:records') {
    const registros = linhas.map(function (linha) { return JSON.parse(linha[1]); });
    valor = JSON.stringify(registros);
  } else {
    // o log é guardado como um único bloco na primeira linha
    valor = linhas[0][1];
  }

  return { ok: true, found: true, value: valor };
}

function gravarChave(chave, valorTexto) {
  const nomeAba = nomeDaAbaParaChave(chave);
  if (!nomeAba) return { ok: false, error: 'Chave desconhecida: ' + chave };

  // Trava para evitar que duas pessoas gravem ao mesmo tempo e
  // um salvamento apague o do outro.
  const trava = LockService.getScriptLock();
  trava.waitLock(30000);

  try {
    const aba = obterOuCriarAba(nomeAba);
    aba.clear();
    aba.appendRow(['id', 'json']);
    aba.setFrozenRows(1);
    aba.getRange(1, 1, 1, 2).setFontWeight('bold');

    if (chave === 'reaberturas:records') {
      const registros = JSON.parse(valorTexto);
      if (registros.length > 0) {
        const linhas = registros.map(function (r) { return [r.id, JSON.stringify(r)]; });
        aba.getRange(2, 1, linhas.length, 2).setValues(linhas);
      }
    } else {
      aba.appendRow(['log', valorTexto]);
    }

    return { ok: true };

  } finally {
    trava.releaseLock();
  }
}


/* ============================================================
 *  GRAVAÇÃO POR REGISTRO (para uso em equipe)
 *  Em vez de reescrever a aba inteira a cada salvamento, estas
 *  funções mexem só na linha do registro em questão. Assim, se
 *  duas pessoas salvarem casos diferentes ao mesmo tempo, uma
 *  não apaga o trabalho da outra.
 * ============================================================ */

function gravarUmRegistro(registro) {
  if (!registro || !registro.id) return { ok: false, error: 'Registro inválido (sem id).' };

  const trava = LockService.getScriptLock();
  trava.waitLock(30000);
  try {
    const aba = obterOuCriarAba(SHEET_BASE);
    const dados = aba.getDataRange().getValues();
    const jsonTexto = JSON.stringify(registro);

    let linhaEncontrada = -1;
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] === registro.id) { linhaEncontrada = i + 1; break; }
    }

    if (linhaEncontrada > 0) {
      aba.getRange(linhaEncontrada, 1, 1, 2).setValues([[registro.id, jsonTexto]]);
    } else {
      aba.appendRow([registro.id, jsonTexto]);
    }

    return { ok: true };
  } finally {
    trava.releaseLock();
  }
}

function removerUmRegistro(id) {
  if (!id) return { ok: false, error: 'id não informado.' };

  const trava = LockService.getScriptLock();
  trava.waitLock(30000);
  try {
    const aba = obterOuCriarAba(SHEET_BASE);
    const dados = aba.getDataRange().getValues();
    for (let i = dados.length - 1; i >= 1; i--) {
      if (dados[i][0] === id) { aba.deleteRow(i + 1); break; }
    }
    return { ok: true };
  } finally {
    trava.releaseLock();
  }
}

function acrescentarLog(texto) {
  const trava = LockService.getScriptLock();
  trava.waitLock(30000);
  try {
    const aba = obterOuCriarAba(SHEET_LOG);
    const dados = aba.getDataRange().getValues();

    let historico = [];
    if (dados.length > 1 && dados[1][1]) {
      try { historico = JSON.parse(dados[1][1]); } catch (erro) { historico = []; }
    }
    historico.unshift({ t: new Date().toISOString(), text: String(texto || '') });
    historico = historico.slice(0, 150);

    aba.clear();
    aba.appendRow(['id', 'json']);
    aba.setFrozenRows(1);
    aba.getRange(1, 1, 1, 2).setFontWeight('bold');
    aba.appendRow(['log', JSON.stringify(historico)]);

    return { ok: true };
  } finally {
    trava.releaseLock();
  }
}


/* ============================================================
 *  FUNÇÃO DE TESTE
 *  Depois de colar este código, clique em cima do nome
 *  "testarConexao" no menu de funções (acima) e depois no
 *  botão ▷ Executar. Se aparecer "OK — sistema pronto" no
 *  registro de execução (Ver > Registro de execução), está
 *  tudo certo.
 * ============================================================ */
function testarConexao() {
  const resultado = gravarChave('reaberturas:log', JSON.stringify([{ t: new Date().toISOString(), text: 'Teste de conexão realizado.' }]));
  if (resultado.ok) {
    Logger.log('OK — sistema pronto. As abas BASE_SISTEMA e LOG_SISTEMA foram criadas.');
  } else {
    Logger.log('Algo deu errado: ' + JSON.stringify(resultado));
  }
}
