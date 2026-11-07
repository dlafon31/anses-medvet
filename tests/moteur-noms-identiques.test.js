const XLSX = require('xlsx');
const R = require('../site/js/refresher.js');
const assert = require('assert');

// ---------- Fausse API Grist en mémoire ----------
function makeDoc() {
  const meta = {
    tables: [
      [2, 'Date_Export_BDD'], [3, 'Medicaments'], [5, 'Substances_Actives'], [10, 'RCP1'], [12, 'Titres_Para_RCP'],
    ],
    cols: [ // id, parent, colId, label, type, isFormula, formula, visibleCol
      [2, 2, 'Date_Export', 'Date Export', 'DateTime:Europe/Paris', 0, '', 0],
      [3, 2, 'Column_1', 'Column_1', 'Text', 0, '', 0],
      [5, 3, 'Num_Ident', 'Num_Ident', 'Numeric', 0, '', 0],
      [6, 3, 'Nom_du_medicament', 'Nom du médicament', 'Text', 0, '', 0],
      [9, 3, 'Date_1ere_AMM', 'Date 1ère AMM', 'Date', 0, '', 0],
      [18, 3, 'Substances_Actives', 'Substances Actives', 'Any', 1, 'Substances_Actives.lookupRecords(...)', 0],
      [27, 5, 'Num_Ident', 'Num_Ident', 'Ref:Medicaments', 0, '', 5],
      [28, 5, 'Substance_active', 'Substance active', 'Text', 0, '', 0],
      [29, 5, 'Quantite', 'Quantité', 'Numeric', 0, '', 0],
      [56, 10, 'N_d_identification', "N° d'identification", 'Ref:Medicaments', 0, '', 5],
      [57, 10, 'N_Para', 'N° Para', 'Ref:Titres_Para_RCP', 0, '', 66],
      [58, 10, 'Contenu_Para', 'Contenu Para', 'Text', 0, '', 0],
      [66, 12, 'N_Para', 'N° Para', 'Numeric', 0, '', 0],
      [67, 12, 'Titre', 'Titre', 'Text', 0, '', 0],
    ],
  };
  const data = {
    Date_Export_BDD: { 1: { Date_Export: 1773760629, Column_1: '' } },
    Medicaments: { 1: { Num_Ident: 916560, Nom_du_medicament: 'CANIXIN L', Date_1ere_AMM: 347673600 },
                   2: { Num_Ident: 974560, Nom_du_medicament: 'ANCIEN NOM', Date_1ere_AMM: 341107200 },
                   3: { Num_Ident: 111, Nom_du_medicament: 'RETIRE', Date_1ere_AMM: null } },
    Substances_Actives: { 1: { Num_Ident: 1, Substance_active: 'X', Quantite: 1 } },
    RCP1: { 1: { N_d_identification: 1, N_Para: 2, Contenu_Para: 'old' } },
    Titres_Para_RCP: { 1: { N_Para: 63, Titre: 'A' }, 2: { N_Para: 70, Titre: 'B' } },
  };
  let nextId = { Date_Export_BDD: 2, Medicaments: 4, Substances_Actives: 2, RCP1: 2, Titres_Para_RCP: 3 };
  const actions = [];
  const colsOf = (t) => meta.cols.filter((c) => c[1] === meta.tables.find((x) => x[1] === t)[0]);
  const docApi = {
    async fetchTable(t) {
      if (t === '_grist_Tables') return { id: meta.tables.map((x) => x[0]), tableId: meta.tables.map((x) => x[1]), summarySourceTable: meta.tables.map(() => 0) };
      if (t === '_grist_Tables_column') {
        const k = ['id', 'parentId', 'colId', 'label', 'type', 'isFormula', 'formula', 'visibleCol'];
        const o = {}; k.forEach((n, j) => { o[n] = meta.cols.map((c) => c[j]); }); return o;
      }
      const rows = data[t]; const ids = Object.keys(rows).map(Number);
      const o = { id: ids };
      colsOf(t).forEach((c) => { o[c[2]] = ids.map((i) => rows[i][c[2]] ?? null); });
      return o;
    },
    async applyUserActions(acts) {
      for (const [kind, t, ids, cols] of acts) {
        actions.push([kind, t, ids.length]);
        if (kind === 'BulkRemoveRecord') ids.forEach((i) => delete data[t][i]);
        else if (kind === 'BulkAddRecord') ids.forEach((_, n) => { const id = nextId[t]++; data[t][id] = {}; for (const c in cols) data[t][id][c] = cols[c][n]; });
        else if (kind === 'BulkUpdateRecord') ids.forEach((id, n) => { for (const c in cols) data[t][id][c] = cols[c][n]; });
      }
    },
  };
  return { docApi, data, actions };
}

// ---------- Classeur de test ----------
function makeWorkbook() {
  const wb = XLSX.utils.book_new();
  const add = (name, aoa) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa, { cellDates: true }), name);
  add('Date_Export_BDD', [['Date Export'], [new Date(Date.UTC(2026, 10, 5, 8, 30))]]);
  add('Médicaments', [
    ['Num_Ident', 'Nom du médicament', 'Date 1ère AMM', 'Nouvelle colonne'],
    [916560, 'CANIXIN L', new Date(Date.UTC(1981, 0, 7)), 'x'],
    [974560, 'CANIGEN CH', '25/10/1980', 'y'],
    ['999001', 'NOUVEAU', null, 'z'],
  ]);
  add('Substances_Actives', [
    ['Num_Ident', 'Substance active', 'Quantité'],
    [916560, 'Leptospira', '1,5'],
    [999001, 'Amoxicilline', 250],
    [123456, 'Orpheline', 3],
  ]);
  add('RCP1', [
    ["N° d'identification", 'N° Para', 'Contenu Para'],
    [916560, 70, 'CANIXIN L'],
    [974560, 63, 'texte'],
  ]);
  add('Feuil_inconnue', [['a'], [1]]);
  // Relecture comme dans le navigateur
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  return XLSX.read(buf, { type: 'buffer' });
}

(async () => {
  // 1) Simulation : aucune écriture
  let { docApi, data, actions } = makeDoc();
  const logs = [];
  let rep = await R.refresh({ docApi, XLSX, workbook: makeWorkbook(), dryRun: true, log: (m) => logs.push(m) });
  assert.strictEqual(actions.length, 0, 'dry-run ne doit rien écrire');
  console.log('DRY-RUN\n' + logs.join('\n'));

  // 2) Exécution réelle
  logs.length = 0;
  rep = await R.refresh({ docApi, XLSX, workbook: makeWorkbook(), log: (m) => logs.push(m) });
  console.log('\nRÉEL\n' + logs.join('\n'));
  console.log(actions);

  const M = data.Medicaments;
  assert.deepStrictEqual(Object.keys(M).map(Number), [1, 2, 4], 'ids stables + nouvel id, 111 supprimé');
  assert.strictEqual(M[2].Nom_du_medicament, 'CANIGEN CH');
  assert.strictEqual(M[1].Date_1ere_AMM, Date.UTC(1981, 0, 7) / 1000);
  assert.strictEqual(M[2].Date_1ere_AMM, Date.UTC(1980, 9, 25) / 1000);
  assert.strictEqual(M[4].Num_Ident, 999001);
  const S = Object.values(data.Substances_Actives);
  assert.deepStrictEqual(S.map((s) => s.Num_Ident), [1, 4, '123456'], 'Ref remappées, orpheline gardée en texte');
  assert.strictEqual(S[0].Quantite, 1.5);
  const P = Object.values(data.RCP1);
  assert.deepStrictEqual(P.map((p) => [p.N_d_identification, p.N_Para]), [[1, 2], [2, 1]], 'Ref vers Titres via N_Para');
  const d = Object.values(data.Date_Export_BDD)[0].Date_Export;
  console.log('Date export →', new Date(d * 1000).toISOString());

  // 3) Idempotence : 2e passage sans aucune écriture
  actions.length = 0; logs.length = 0;
  rep = await R.refresh({ docApi, XLSX, workbook: makeWorkbook(), log: (m) => logs.push(m) });
  console.log('\n2e PASSAGE\n' + logs.join('\n'));
  assert.strictEqual(actions.length, 0, 'second passage : aucune écriture attendue');
  console.log('\nTOUS LES TESTS PASSENT');
})().catch((e) => { console.error(e); process.exit(1); });
