/**
 * Backend de escrita para a página de preenchimento (entrada.html).
 *
 * Como instalar:
 * 1. Abra a planilha "Sistema de gestão" no Google Sheets.
 * 2. Menu Extensões -> Apps Script.
 * 3. Apague o conteúdo padrão do arquivo Code.gs e cole este arquivo inteiro.
 * 4. Clique em Implantar -> Nova implantação -> tipo "App da Web".
 *    - Executar como: Eu (sua conta)
 *    - Quem pode acessar: Qualquer pessoa
 * 5. Autorize as permissões pedidas (é a sua própria planilha).
 * 6. Copie a URL do app da Web gerada e cole em entrada.html na constante SCRIPT_URL.
 *
 * Sempre que editar este script depois de já implantado, use
 * "Implantar -> Gerenciar implantações -> Editar -> Nova versão" para as
 * mudanças valerem (só salvar o arquivo não atualiza a URL já publicada).
 *
 * Funciona tanto colado dentro da planilha (Extensões -> Apps Script) quanto
 * num projeto avulso do Apps Script — nos dois casos ele abre a planilha
 * pelo ID abaixo em vez de depender de "planilha ativa".
 */
var SPREADSHEET_ID = '1XBswfpypHskEIG75ZrVm5_Gu7A8FaVbs';
function getSS() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

// Datas voltam como objeto Date do Sheets; convertidas pra "AAAA-MM-DD" (o
// formato que <input type=date> espera) usando o calendário local, não UTC.
function dateToIso(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    var y = v.getFullYear(), m = v.getMonth() + 1, d = v.getDate();
    return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  }
  return String(v);
}

function doGet(e) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName('Banco de atletas');
    var col = headerMap(sheet, 1);
    var lastRow = sheet.getLastRow();
    var rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];
    var athletes = rows.map(function (r) {
      return {
        id: r[(col['ID atleta'] || 3) - 1],
        nome: r[(col['Nome'] || 1) - 1],
        apelido: r[(col['Apelido'] || 2) - 1],
        dataNascimento: dateToIso(r[(col['Data de Nascimento'] || 4) - 1]),
        posicao: r[(col['Posição'] || 6) - 1],
        quartil: r[(col['Quartil'] || 5) - 1],
        peDominante: r[(col['Pé dominante'] || 7) - 1],
        chegadaAoClube: dateToIso(r[(col['Chegada ao clube'] || 8) - 1]),
        categoria: r[(col['Categoria'] || 9) - 1],
        foto: r[(col['Link foto'] || 10) - 1],
        caracteristica: r[(col['Característica'] || 11) - 1],
        pontosAMelhorar: r[(col['Pontos a melhorar'] || 12) - 1],
      };
    }).filter(function (a) { return a.apelido; });
    return jsonOut({ ok: true, athletes: athletes });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var result;
    if (action === 'addAthlete') result = addAthlete(body.data);
    else if (action === 'updateAthlete') result = updateAthlete(body.id, body.data);
    else if (action === 'addTraining') result = addTraining(body.data);
    else if (action === 'addAvaliacoes') result = addAvaliacoes(body.date, body.entries);
    else if (action === 'addGame') result = addGame(body.data);
    else if (action === 'addGameInd') result = addGameInd(body.gameNo, body.date, body.category, body.entries);
    else throw new Error('Ação desconhecida: ' + action);
    return jsonOut({ ok: true, result: result });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Mapeia texto do cabeçalho (normalizado) -> número da coluna (1-based),
// pra não depender da ordem/posição exata das colunas na planilha.
function headerMap(sheet, headerRow) {
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var map = {};
  values.forEach(function (v, i) {
    if (v) {
      var key = String(v).replace(/\s+/g, ' ').trim();
      if (!(key in map)) map[key] = i + 1;
    }
  });
  return map;
}

function setByHeader(sheet, row, colmap, name, value) {
  var col = colmap[name];
  if (!col) return;
  if (value === undefined || value === null || value === '') return;
  sheet.getRange(row, col).setValue(value);
}

// Datas chegam como texto "AAAA-MM-DD" (do <input type=date>); convertidas
// pra objeto Date real aqui pra caírem como célula de data de verdade na
// planilha, e não como texto.
function setDateByHeader(sheet, row, colmap, name, isoStr) {
  var col = colmap[name];
  if (!col || !isoStr) return;
  var parts = String(isoStr).split('-');
  if (parts.length !== 3) return;
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  sheet.getRange(row, col).setValue(d);
}

// Copia a fórmula da última linha preenchida para a linha nova, pras colunas
// calculadas (TAA, % de minutagem, Saldo, Resultado etc.) continuarem
// funcionando como no resto da planilha.
function copyFormulasFromAbove(sheet, aboveRow, newRow, colmap, names) {
  if (aboveRow < 1) return;
  names.forEach(function (name) {
    var c = colmap[name];
    if (!c) return;
    var formula = sheet.getRange(aboveRow, c).getFormula();
    if (formula) {
      var re = new RegExp(aboveRow, 'g');
      sheet.getRange(newRow, c).setFormula(formula.replace(re, newRow));
    }
  });
}

function nextNumber(sheet, startRow, col, count) {
  if (count <= 0) return 1;
  var vals = sheet.getRange(startRow, col, count, 1).getValues().flat()
    .filter(function (v) { return v !== '' && v !== null; })
    .map(Number).filter(function (v) { return !isNaN(v); });
  return vals.length ? Math.max.apply(null, vals) + 1 : 1;
}

// ---------- Banco de atletas ----------
function addAthlete(d) {
  var sheet = getSS().getSheetByName('Banco de atletas');
  var col = headerMap(sheet, 1);
  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;
  var nextId = nextNumber(sheet, 2, col['ID atleta'], Math.max(lastRow - 1, 0));
  setByHeader(sheet, newRow, col, 'Nome', d.nome);
  setByHeader(sheet, newRow, col, 'Apelido', d.apelido);
  setByHeader(sheet, newRow, col, 'ID atleta', nextId);
  setDateByHeader(sheet, newRow, col, 'Data de Nascimento', d.dataNascimento);
  setByHeader(sheet, newRow, col, 'Quartil', d.quartil);
  setByHeader(sheet, newRow, col, 'Posição', d.posicao);
  setByHeader(sheet, newRow, col, 'Pé dominante', d.peDominante);
  setDateByHeader(sheet, newRow, col, 'Chegada ao clube', d.chegadaAoClube);
  setByHeader(sheet, newRow, col, 'Categoria', d.categoria || 'Sub-10');
  setByHeader(sheet, newRow, col, 'Link foto', d.linkFoto);
  setByHeader(sheet, newRow, col, 'Característica', d.caracteristica);
  setByHeader(sheet, newRow, col, 'Pontos a melhorar', d.pontosAMelhorar);
  return { id: nextId, row: newRow };
}

// Atualiza um atleta já cadastrado (localizado pelo ID atleta). Campos em
// branco no formulário não sobrescrevem o que já está na planilha (mesma
// regra de setByHeader usada no cadastro novo).
function updateAthlete(id, d) {
  var sheet = getSS().getSheetByName('Banco de atletas');
  var col = headerMap(sheet, 1);
  var lastRow = sheet.getLastRow();
  var idCol = col['ID atleta'];
  var ids = lastRow > 1 ? sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat() : [];
  var rowIdx = ids.findIndex(function (v) { return String(v) === String(id); });
  if (rowIdx === -1) throw new Error('Atleta não encontrado (ID ' + id + ').');
  var row = rowIdx + 2;
  setByHeader(sheet, row, col, 'Nome', d.nome);
  setByHeader(sheet, row, col, 'Apelido', d.apelido);
  setDateByHeader(sheet, row, col, 'Data de Nascimento', d.dataNascimento);
  setByHeader(sheet, row, col, 'Quartil', d.quartil);
  setByHeader(sheet, row, col, 'Posição', d.posicao);
  setByHeader(sheet, row, col, 'Pé dominante', d.peDominante);
  setDateByHeader(sheet, row, col, 'Chegada ao clube', d.chegadaAoClube);
  setByHeader(sheet, row, col, 'Categoria', d.categoria);
  setByHeader(sheet, row, col, 'Link foto', d.linkFoto);
  setByHeader(sheet, row, col, 'Característica', d.caracteristica);
  setByHeader(sheet, row, col, 'Pontos a melhorar', d.pontosAMelhorar);
  return { id: id, row: row };
}

// ---------- Banco de treinos ----------
function addTraining(d) {
  var sheet = getSS().getSheetByName('Banco de treinos');
  var headerRow = 3;
  var col = headerMap(sheet, headerRow);
  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;

  var fields = {
    'ID da sessão': d.idSessao, 'Categoria': d.categoria || 'Sub-10',
    'Nota geral da sessão\n(0 a 10)': d.notaGeral, 'Nº de atletas': d.numAtletas,
    'Atletas em avaliação': d.atletasAvaliacao, 'Observações': d.observacoes,
  };
  ['preparatoria', 'conceitual'].forEach(function (key) {
    var suf = key === 'preparatoria' ? 'preparatória' : 'conceitual';
    var p = d[key] || {};
    fields['Conteúdo principal - ' + suf] = p.conteudoPrincipal;
    fields['Conteúdo secundário - ' + suf] = p.conteudoSecundario;
    fields['Conteúdo concorrente - ' + suf] = p.conteudoConcorrente;
    fields['Tipo de conteúdo - ' + suf] = p.tipoConteudo;
    fields['Tipo de tarefa - ' + suf] = p.tipoTarefa;
    fields['Complexidade - ' + suf] = p.complexidade;
    fields['Minutagem - ' + suf] = p.minutagem;
    fields['Jogadores ativos - ' + suf] = p.jogadoresAtivos;
    fields['Desempenho - ' + suf + ' (0 a 10)'] = p.desempenho;
  });
  var conexao = d.conexao || {};
  fields['Relação númerica - conexão'] = conexao.relacaoNumerica;
  fields['Tipo de tarefa - conexão'] = conexao.tipoTarefa;
  fields['Complexidade - conexão'] = conexao.complexidade;
  fields['Minutagem - conexão'] = conexao.minutagem;
  fields['Jogadores ativos - conexão'] = conexao.jogadoresAtivos;
  fields['Desempenho - conexão (0 a 10)'] = conexao.desempenho;

  setDateByHeader(sheet, newRow, col, 'Data', d.data);
  Object.keys(fields).forEach(function (name) { setByHeader(sheet, newRow, col, name, fields[name]); });

  var formulaCols = ['TAA - preparatória', 'TAA - conceitual', 'TAA - conexão',
    'Minutagem efetiva', 'Total TAA', '% de minutagem', '% de TAA'];
  copyFormulasFromAbove(sheet, lastRow, newRow, col, formulaCols);
  return { row: newRow };
}

// ---------- Avaliação individual do treino ----------
function addAvaliacoes(dateStr, entries) {
  var sheet = getSS().getSheetByName('Avaliação individual do treino');
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var dateCol = null;
  for (var i = 2; i < headers.length; i++) {
    if (String(headers[i]).trim() === dateStr) { dateCol = i + 1; break; }
  }
  if (!dateCol) {
    dateCol = lastCol + 1;
    sheet.getRange(1, dateCol).setValue(dateStr);
  }
  var lastRow = sheet.getLastRow();
  var nicknames = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  var written = 0;
  (entries || []).forEach(function (en) {
    var rowIdx = nicknames.indexOf(en.apelido);
    if (rowIdx === -1) return;
    if (en.nota === '' || en.nota === null || en.nota === undefined) return;
    sheet.getRange(rowIdx + 2, dateCol).setValue(Number(en.nota));
    written++;
  });
  return { dateCol: dateCol, written: written };
}

// ---------- Pós-jogo ----------
function addGame(d) {
  var sheet = getSS().getSheetByName('Pós-jogo');
  var headerRow = 2;
  var col = headerMap(sheet, headerRow);
  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;
  var gameNo = d.numeroJogo || nextNumber(sheet, headerRow + 1, col['Nº Jogo'], Math.max(lastRow - headerRow, 0));

  var fields = {
    'Nº Jogo': gameNo, 'Categoria': d.categoria || 'Sub-10', 'Adversário': d.adversario,
    'Ranking': d.ranking, 'Categoria adversário': d.categoriaAdversario, 'Competição': d.competicao,
    'Local': d.local, 'Cidade': d.cidade, 'Período': d.periodo, 'Formato': d.formato,
    'Plataforma de jogo': d.plataforma, 'Treinador': d.treinador, 'Capitão': d.capitao,
    'Minutagem': d.minutagem, 'Gols feitos': d.golsFeitos, 'Gols sofridos': d.golsSofridos,
    'Cartão amarelo': d.cartaoAmarelo, 'Cartão vermelho': d.cartaoVermelho,
    'Desempenho': d.desempenho, 'Comentários gerais': d.comentariosGerais,
    'Pontos positivos': d.pontosPositivos, 'Pontos negativos': d.pontosNegativos,
  };
  setDateByHeader(sheet, newRow, col, 'Data', d.data);
  Object.keys(fields).forEach(function (name) { setByHeader(sheet, newRow, col, name, fields[name]); });

  var formulaCols = ['Saldo', 'Resultado', 'Data entre jogos'];
  copyFormulasFromAbove(sheet, lastRow, newRow, col, formulaCols);
  return { gameNo: gameNo, row: newRow };
}

// ---------- Pós-jogo individual ----------
function addGameInd(gameNo, dateStr, categoria, entries) {
  var sheet = getSS().getSheetByName('Pós-jogo individual');
  var headerRow = 2;
  var col = headerMap(sheet, headerRow);
  var startRow = sheet.getLastRow() + 1;
  (entries || []).forEach(function (en, idx) {
    var row = startRow + idx;
    var fields = {
      'Nº Jogo': gameNo, 'Categoria': categoria || 'Sub-10', 'Nome': en.apelido,
      'Posição': en.posicao, 'Titularidade': en.titularidade, 'Minutagem': en.minutagem,
      'Desempenho': en.desempenho, 'Gols feitos': en.golsFeitos, 'Gols sofridos': en.golsSofridos,
      'Assistências': en.assistencias, 'Cartão amarelo': en.cartaoAmarelo, 'Cartão vermelho': en.cartaoVermelho,
      'Comentários individuais': en.comentarios,
    };
    setDateByHeader(sheet, row, col, 'Data', dateStr);
    Object.keys(fields).forEach(function (name) { setByHeader(sheet, row, col, name, fields[name]); });
  });
  return { rows: (entries || []).length };
}
