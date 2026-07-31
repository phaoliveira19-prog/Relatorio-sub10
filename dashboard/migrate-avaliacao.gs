/**
 * Migração ÚNICA da aba "Avaliação individual do treino" do formato largo
 * (uma coluna nova por data de treino) pro formato longo (uma linha por
 * atleta-por-treino: Data | Apelido | Nota).
 *
 * SEGURO: este script só LÊ a aba antiga e CRIA uma aba nova — não apaga
 * nem altera nada na aba original. Só depois de conferir que a aba nova
 * ficou certa é que você renomeia as abas manualmente (passo 4 abaixo).
 *
 * Como rodar:
 * 1. Abra a planilha -> Extensões -> Apps Script.
 * 2. Crie um arquivo novo (ícone "+" ao lado de "Arquivos" -> Script) chamado
 *    "migrate-avaliacao" e cole este conteúdo (pode ficar junto com o
 *    Code.gs do apps-script.gs, não precisa ser em arquivo separado, mas
 *    fica mais organizado).
 * 3. No seletor de função (barra de cima, ao lado do ícone de "Executar"),
 *    escolha "migrateAvaliacaoIndividual" e clique em "Executar".
 *    Na primeira vez vai pedir autorização — é a sua própria planilha, pode
 *    aceitar.
 * 4. Confira a aba nova "Avaliação individual (novo formato)" que apareceu.
 *    Se estiver tudo certo:
 *    a) Clique com o botão direito na aba antiga "Avaliação individual do
 *       treino" -> Renomear -> algo como "Avaliação individual (arquivo)".
 *    b) Renomeie "Avaliação individual (novo formato)" para
 *       "Avaliação individual do treino" (o nome exato que o site espera).
 * 5. Me avise quando terminar — nesse momento eu aviso quando o site e a
 *    página de preenchimento já estiverem lendo/escrevendo no formato novo.
 */
function migrateAvaliacaoIndividual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var oldSheet = ss.getSheetByName('Avaliação individual do treino');
  if (!oldSheet) throw new Error('Aba "Avaliação individual do treino" não encontrada.');

  var lastCol = oldSheet.getLastColumn();
  var headerRow = oldSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var dateCols = [];
  for (var i = 2; i < headerRow.length; i++) {
    if (headerRow[i]) dateCols.push({ idx: i, label: String(headerRow[i]).trim() });
  }

  var lastRow = oldSheet.getLastRow();
  var data = lastRow > 1 ? oldSheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  var newName = 'Avaliação individual (novo formato)';
  var existing = ss.getSheetByName(newName);
  if (existing) ss.deleteSheet(existing); // permite rodar de novo do zero se precisar refazer
  var newSheet = ss.insertSheet(newName);
  newSheet.getRange(1, 1, 1, 3).setValues([['Data', 'Apelido', 'Nota']]);

  var out = [];
  data.forEach(function (row) {
    var nickname = row[0];
    if (!nickname) return;
    dateCols.forEach(function (dc) {
      var note = row[dc.idx];
      if (note === '' || note === null || note === undefined) return;
      var parts = dc.label.split('/'); // formato do cabeçalho: DD/MM/AAAA
      if (parts.length !== 3) return;
      var dt = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
      out.push([dt, nickname, note]);
    });
  });

  if (out.length) {
    newSheet.getRange(2, 1, out.length, 3).setValues(out);
    newSheet.getRange(2, 1, out.length, 1).setNumberFormat('dd/mm/yyyy');
  }
  newSheet.autoResizeColumns(1, 3);
  Logger.log('Migradas ' + out.length + ' notas para a aba "' + newName + '".');
}
