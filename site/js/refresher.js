/* ==========================================================================
 * Moteur d'actualisation : classeur Excel  ->  tables Grist
 * Aucune dépendance au DOM (testable hors navigateur).
 * ========================================================================== */
const Refresher = (() => {
  'use strict';

  // ---------- Utilitaires ----------
  const norm = (s) => String(s ?? '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');

  const isBlank = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');

  function tzOffsetMs(tz, utcMs) {
    const dtf = new Intl.DateTimeFormat('en-US', {
      timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    });
    const p = Object.fromEntries(dtf.formatToParts(new Date(utcMs)).map((x) => [x.type, x.value]));
    return Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second) - utcMs;
  }

  function localPartsToEpoch(parts, tz) {
    const guess = Date.UTC(...parts);
    if (!tz) return guess / 1000;
    const off = tzOffsetMs(tz, guess);
    let t = guess - off;
    const off2 = tzOffsetMs(tz, t);
    if (off2 !== off) t = guess - off2;
    return t / 1000;
  }

  // Renvoie [Y, M(0-11), D, h, m, s], null si vide, undefined si illisible
  function dateParts(v) {
    if (isBlank(v)) return null;
    if (v instanceof Date) return [v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds()];
    if (typeof v === 'number') {
      const d = new Date(Math.round((v - 25569) * 86400) * 1000); // date série Excel
      return [d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()];
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) return [+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)];
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) return [+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)];
    m = s.match(/(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})(?:\D{1,4}(\d{1,2})[:hH](\d{2})(?::(\d{2}))?)?/);
    if (m) return [+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)];
    return undefined;
  }

  function toNumber(v) {
    if (isBlank(v)) return null;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    const s = String(v).replace(/[\s ]/g, '').replace(',', '.');
    if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return Number(s);
    return String(v); // conservé tel quel (Grist l'affichera comme valeur invalide)
  }

  // Clé de correspondance pour les références (916560 == "916560" == 916560.0)
  function keyOf(v) {
    if (isBlank(v)) return null;
    if (typeof v === 'number') return String(v);
    const s = String(v).trim();
    const n = toNumber(s);
    return typeof n === 'number' ? String(n) : s.toLowerCase();
  }

  // Conversion d'une cellule Excel vers le type de la colonne Grist
  function convert(v, col, refMaps, stats) {
    const type = col.type || 'Any';
    if (type === 'Text' || type.startsWith('Choice')) {
      if (isBlank(v)) return '';
      if (v instanceof Date) return v.toISOString();
      return String(v);
    }
    if (type === 'Numeric' || type === 'Int' || type === 'ManualSortPos') {
      const n = toNumber(v);
      return type === 'Int' && typeof n === 'number' ? Math.round(n) : n;
    }
    if (type === 'Bool') {
      if (isBlank(v)) return false;
      if (typeof v === 'boolean') return v;
      return ['1', 'true', 'vrai', 'oui', 'yes', 'x'].includes(String(v).trim().toLowerCase());
    }
    if (type === 'Date') {
      const p = dateParts(v);
      if (p === null) return null;
      if (p === undefined) { stats.badDates++; return String(v); }
      return Date.UTC(p[0], p[1], p[2]) / 1000;
    }
    if (type.startsWith('DateTime')) {
      const p = dateParts(v);
      if (p === null) return null;
      if (p === undefined) { stats.badDates++; return String(v); }
      const tz = type.split(':')[1] || null;
      return localPartsToEpoch(p, tz);
    }
    if (type.startsWith('Ref:')) {
      const k = keyOf(v);
      if (k === null) return 0;
      const map = refMaps[col.colId];
      const id = map && map.get(k);
      if (id) return id;
      stats.badRefs++;
      return isBlank(v) ? 0 : String(v); // valeur conservée, visible en rouge dans Grist
    }
    // Any / autres types
    return isBlank(v) ? null : v;
  }

  // Égalité tolérante : '' / null / undefined (et 0 pour une référence vide) sont équivalents
  const nul = (v) => v === '' || v === null || v === undefined || v === 0;
  const same = (a, b) => {
    if (nul(a) && nul(b)) return true;
    if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a));
    return a === b || JSON.stringify(a) === JSON.stringify(b);
  };

  // ---------- Lecture du schéma Grist ----------
  async function readSchema(docApi) {
    const T = await docApi.fetchTable('_grist_Tables');
    const C = await docApi.fetchTable('_grist_Tables_column');
    const byRef = new Map();
    T.id.forEach((ref, i) => {
      if (T.summarySourceTable[i]) return;
      byRef.set(ref, { tableId: T.tableId[i], columns: [] });
    });
    const colById = new Map();
    C.id.forEach((id, i) => colById.set(id, C.colId[i]));
    const order = C.id.map((_, i) => i).sort((a, b) => (C.parentPos ? C.parentPos[a] - C.parentPos[b] : 0) || a - b);
    order.forEach((i) => {
      const t = byRef.get(C.parentId[i]);
      if (!t) return;
      const colId = C.colId[i];
      if (colId === 'manualSort' || colId.startsWith('gristHelper_')) return;
      t.columns.push({
        colId,
        label: C.label[i],
        type: C.type[i],
        isData: !C.isFormula[i] || !String(C.formula[i] || '').trim(),
        visibleColId: C.visibleCol[i] ? colById.get(C.visibleCol[i]) : null,
      });
    });
    const tables = {};
    for (const t of byRef.values()) tables[t.tableId] = t;
    return tables;
  }

  // ---------- Analyse du classeur ----------
  // Correspondance onglet -> table lue dans la table Admin (colonnes Table et Onglet)
  async function readSheetMap(docApi, tables, log = () => {}) {
    const map = new Map(); // norm(onglet) -> tableId
    let admin;
    try { admin = await docApi.fetchTable('Admin'); } catch (e) { return map; }
    if (!admin || !admin.Table || !admin.Onglet) return map;
    const byNorm = new Map(Object.keys(tables).map((t) => [norm(t), t]));
    admin.Table.forEach((t, i) => {
      const onglet = admin.Onglet[i];
      if (isBlank(t) || isBlank(onglet)) return;
      const tableId = byNorm.get(norm(t));
      if (!tableId) { log(`Admin : table « ${String(t).trim()} » introuvable dans le document (ligne ignorée)`, 'warn'); return; }
      map.set(norm(onglet), tableId);
    });
    return map;
  }

  // Correspondance en-tête Excel -> colonne Grist lue dans la table Admin_Colonnes
  // (colonnes Table, Colonne, Entete). Renvoie Map(tableId -> Map(norm(entête) -> colId)).
  async function readColumnAliases(docApi, tables, log = () => {}) {
    const out = new Map();
    let a;
    try { a = await docApi.fetchTable('Admin_Colonnes'); } catch (e) { return out; }
    if (!a || !a.Table || !a.Colonne || !a.Entete) return out;
    const byNorm = new Map(Object.keys(tables).map((t) => [norm(t), t]));
    a.Table.forEach((t, i) => {
      if (isBlank(t) || isBlank(a.Colonne[i]) || isBlank(a.Entete[i])) return;
      const tableId = byNorm.get(norm(t));
      const col = tableId && tables[tableId].columns.find((c) => norm(c.colId) === norm(a.Colonne[i]) || norm(c.label) === norm(a.Colonne[i]));
      if (!col) { log(`Admin_Colonnes : colonne « ${String(t).trim()}.${String(a.Colonne[i]).trim()} » introuvable (ligne ignorée)`, 'warn'); return; }
      if (!out.has(tableId)) out.set(tableId, new Map());
      out.get(tableId).set(norm(a.Entete[i]), col.colId);
    });
    return out;
  }

  function analyseWorkbook(workbook, XLSX, tables, sheetMap = new Map(), colAliases = new Map()) {
    const byNorm = new Map(Object.keys(tables).map((t) => [norm(t), t]));
    const plans = [];
    const ignoredSheets = [];
    const taken = new Set();
    for (const sheetName of workbook.SheetNames) {
      // 1) correspondance déclarée dans Admin, 2) à défaut, même nom que la table
      let tableId = sheetMap.get(norm(sheetName)) || byNorm.get(norm(sheetName));
      if (tableId && taken.has(tableId)) tableId = null; // une seule source par table
      if (tableId) taken.add(tableId);
      if (!tableId) { ignoredSheets.push(sheetName); continue; }
      const ws = workbook.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
      const hIdx = aoa.findIndex((r) => r && r.some((c) => !isBlank(c)));
      const header = hIdx >= 0 ? aoa[hIdx] : [];
      const dataCols = tables[tableId].columns.filter((c) => c.isData);
      const used = new Set();
      const mapping = []; // {idx, col}
      const unmatchedHeaders = [];
      header.forEach((h, idx) => {
        if (isBlank(h)) return;
        const n = norm(h);
        const alias = colAliases.get(tableId) && colAliases.get(tableId).get(n);
        const col = (alias && dataCols.find((c) => !used.has(c.colId) && c.colId === alias))
                 || dataCols.find((c) => !used.has(c.colId) && norm(c.label) === n)
                 || dataCols.find((c) => !used.has(c.colId) && norm(c.colId) === n);
        if (col) { used.add(col.colId); mapping.push({ idx, col }); } else unmatchedHeaders.push(String(h));
      });
      // Aucun en-tête reconnu : l'onglet n'a pas de ligne d'en-tête (ex. « Date export »).
      // La 1re ligne est alors une ligne de données, associée aux colonnes Grist par position.
      let headerless = false;
      if (!mapping.length && hIdx >= 0) {
        headerless = true;
        unmatchedHeaders.length = 0;
        const width = Math.max(...aoa.slice(hIdx).map((r) => r.length));
        for (let idx = 0; idx < Math.min(width, dataCols.length); idx++) {
          // On ne prend que les colonnes qui contiennent au moins une valeur
          if (aoa.slice(hIdx).some((r) => !isBlank(r[idx]))) { mapping.push({ idx, col: dataCols[idx] }); used.add(dataCols[idx].colId); }
        }
      }
      const unmappedCols = dataCols.filter((c) => !used.has(c.colId)).map((c) => c.colId);
      const first = headerless ? hIdx : hIdx + 1;
      const rows = hIdx >= 0 ? aoa.slice(first).filter((r) => r && r.some((c) => !isBlank(c))) : [];
      plans.push({ sheetName, tableId, mapping, unmatchedHeaders, unmappedCols, rows, headerless });
    }
    return { plans, ignoredSheets };
  }

  // Tri topologique : tables référencées avant celles qui les référencent
  function orderPlans(plans) {
    const ids = new Set(plans.map((p) => p.tableId));
    const deps = new Map(plans.map((p) => [p.tableId, new Set(
      p.mapping.map((m) => m.col.type).filter((t) => t.startsWith('Ref:'))
        .map((t) => t.slice(4)).filter((t) => ids.has(t) && t !== p.tableId),
    )]));
    const out = []; const seen = new Set();
    const visit = (t, stack = new Set()) => {
      if (seen.has(t)) return;
      if (stack.has(t)) return; // cycle : on ignore
      stack.add(t);
      for (const d of deps.get(t) || []) visit(d, stack);
      seen.add(t); out.push(plans.find((p) => p.tableId === t));
    };
    plans.forEach((p) => visit(p.tableId));
    return out;
  }

  // Colonne clé d'une table "parente" = colonne affichée par les références qui pointent vers elle
  function keyColumnFor(tableId, tables) {
    for (const t of Object.values(tables)) {
      for (const c of t.columns) {
        if (c.type === 'Ref:' + tableId && c.visibleColId) return c.visibleColId;
      }
    }
    return null;
  }

  // ---------- Écriture par lots ----------
  const MAX_BYTES = 2_500_000;
  const MAX_ROWS = 5000;

  async function bulkWrite(docApi, kind, tableId, ids, colIds, rows, onChunk) {
    let start = 0;
    while (start < rows.length) {
      let end = start; let bytes = 0;
      while (end < rows.length && end - start < MAX_ROWS) {
        bytes += JSON.stringify(rows[end]).length;
        if (bytes > MAX_BYTES && end > start) break;
        end++;
      }
      const cols = {};
      colIds.forEach((c, j) => { cols[c] = rows.slice(start, end).map((r) => r[j]); });
      const chunkIds = kind === 'BulkAddRecord' ? new Array(end - start).fill(null) : ids.slice(start, end);
      await docApi.applyUserActions([[kind, tableId, chunkIds, cols]]);
      start = end;
      onChunk && onChunk(start, rows.length);
    }
  }

  async function bulkRemove(docApi, tableId, ids, onChunk) {
    for (let i = 0; i < ids.length; i += 10000) {
      await docApi.applyUserActions([['BulkRemoveRecord', tableId, ids.slice(i, i + 10000)]]);
      onChunk && onChunk(Math.min(i + 10000, ids.length), ids.length);
    }
  }

  // ---------- Actualisation ----------
  async function refresh({ docApi, XLSX, workbook, dryRun = false, log = () => {}, progress = () => {} }) {
    const tables = await readSchema(docApi);
    const sheetMap = await readSheetMap(docApi, tables, log);
    if (sheetMap.size) {
      log(`Correspondances onglet → table lues dans Admin : ${sheetMap.size}`);
      const sheets = new Set(workbook.SheetNames.map(norm));
      const missing = [...sheetMap.entries()].filter(([o]) => !sheets.has(o)).map(([, t]) => t);
      if (missing.length) log(`Onglets déclarés dans Admin mais absents du fichier, tables non actualisées : ${missing.join(', ')}`, 'warn');
    }
    const colAliases = await readColumnAliases(docApi, tables, log);
    if (colAliases.size) log(`Correspondances de colonnes lues dans Admin_Colonnes : ${[...colAliases.values()].reduce((n, m) => n + m.size, 0)}`);
    const { plans, ignoredSheets } = analyseWorkbook(workbook, XLSX, tables, sheetMap, colAliases);
    const ordered = orderPlans(plans);
    const report = { dryRun, tables: [], ignoredSheets, untouchedTables: [] };

    const plannedIds = new Set(plans.map((p) => p.tableId));
    report.untouchedTables = Object.keys(tables).filter((t) => !plannedIds.has(t));
    if (ignoredSheets.length) log(`Onglets sans table correspondante (ignorés) : ${ignoredSheets.join(', ')}`, 'warn');

    // État final simulé (dry-run) ou relu (réel) de chaque table, pour construire les correspondances de Ref
    const finalState = {};
    const getState = async (tableId) => finalState[tableId] || (finalState[tableId] = await docApi.fetchTable(tableId));
    const buildRefMaps = async (plan) => {
      const maps = {};
      for (const { col } of plan.mapping) {
        if (!col.type.startsWith('Ref:')) continue;
        const target = col.type.slice(4);
        if (!tables[target]) continue;
        const keyCol = col.visibleColId;
        if (!keyCol) { maps[col.colId] = null; continue; }
        const st = await getState(target);
        const m = new Map();
        (st.id || []).forEach((id, i) => { const k = keyOf(st[keyCol] && st[keyCol][i]); if (k !== null && !m.has(k)) m.set(k, id); });
        maps[col.colId] = m;
      }
      return maps;
    };

    let step = 0;
    for (const plan of ordered) {
      step++;
      const { tableId, mapping } = plan;
      const r = { tableId, sheet: plan.sheetName, rowsExcel: plan.rows.length, added: 0, updated: 0, removed: 0,
        mode: '', badRefs: 0, badDates: 0, unmatchedHeaders: plan.unmatchedHeaders, unmappedCols: plan.unmappedCols, skipped: false };
      progress({ step, total: ordered.length, tableId, phase: 'analyse' });
      log(`— ${tableId} : ${plan.rows.length} lignes dans l'onglet « ${plan.sheetName} »`);
      if (plan.headerless) log(`  Onglet sans ligne d'en-tête : colonnes associées par position (${mapping.map((m) => m.col.colId).join(', ')})`);
      if (plan.unmatchedHeaders.length) log(`  Colonnes Excel sans équivalent dans Grist : ${plan.unmatchedHeaders.join(', ')}`, 'warn');
      if (plan.unmappedCols.length) log(`  Colonnes Grist absentes de l'Excel : ${plan.unmappedCols.join(', ')}`, 'warn');
      if (!mapping.length) { log('  Aucune colonne commune, table ignorée.', 'error'); r.skipped = true; report.tables.push(r); continue; }

      const parentKey = keyColumnFor(tableId, tables);
      if (parentKey && tables[tableId].columns.some((c) => c.colId === parentKey && c.isData)
          && !mapping.some((m) => m.col.colId === parentKey)) {
        log(`  ARRÊT : la colonne clé « ${parentKey} » (utilisée par les références) n'a pas été trouvée dans l'onglet. `
          + `Table non modifiée. Déclarez l'en-tête correspondant dans Admin_Colonnes.`, 'error');
        r.mode = `ignorée : clé ${parentKey} absente`; r.skipped = true; r.blocked = true;
        report.tables.push(r); continue;
      }
      const missingRefs = tables[tableId].columns
        .filter((c) => c.isData && c.type.startsWith('Ref:') && !mapping.some((m) => m.col.colId === c.colId))
        .map((c) => c.colId);
      if (missingRefs.length) {
        log(`  ARRÊT : colonne(s) de référence ${missingRefs.join(', ')} introuvable(s) dans l'onglet. `
          + `Table non modifiée. Déclarez l'en-tête correspondant dans Admin_Colonnes.`, 'error');
        r.mode = `ignorée : ${missingRefs.join(', ')} absente(s)`; r.skipped = true; r.blocked = true;
        report.tables.push(r); continue;
      }
      const refMaps = await buildRefMaps(plan);
      const stats = { badRefs: 0, badDates: 0 };
      const colIds = mapping.map((m) => m.col.colId);
      const newRows = plan.rows.map((row) => mapping.map(({ idx, col }) => convert(row[idx], col, refMaps, stats)));
      r.badRefs = stats.badRefs; r.badDates = stats.badDates;
      if (stats.badRefs) log(`  ${stats.badRefs} référence(s) sans correspondance (conservées comme texte)`, 'warn');
      if (stats.badDates) log(`  ${stats.badDates} date(s) illisible(s) (conservées comme texte)`, 'warn');

      const existing = await docApi.fetchTable(tableId);
      const exIds = existing.id || [];
      const keyCol = keyColumnFor(tableId, tables);
      const keyPos = keyCol ? colIds.indexOf(keyCol) : -1;
      const newKeys = keyPos >= 0 ? newRows.map((row) => keyOf(row[keyPos])) : null;
      const canUpsert = keyPos >= 0 && newKeys.every((k) => k !== null) && new Set(newKeys).size === newKeys.length;

      if (canUpsert) {
        // Mise à jour ligne à ligne : les identifiants de ligne restent stables
        r.mode = `mise à jour par clé (${keyCol})`;
        const exByKey = new Map();
        const dupIds = [];
        exIds.forEach((id, i) => {
          const k = keyOf(existing[keyCol][i]);
          if (k === null || exByKey.has(k)) dupIds.push(id); else exByKey.set(k, i);
        });
        const toAdd = []; const updIds = []; const updRows = []; const seen = new Set();
        newRows.forEach((row, n) => {
          const k = newKeys[n];
          seen.add(k);
          if (!exByKey.has(k)) { toAdd.push(row); return; }
          const i = exByKey.get(k);
          if (colIds.some((c, j) => !same(existing[c] ? existing[c][i] : null, row[j]))) { updIds.push(exIds[i]); updRows.push(row); }
        });
        const toRemove = dupIds.concat([...exByKey.entries()].filter(([k]) => !seen.has(k)).map(([, i]) => exIds[i]));
        r.added = toAdd.length; r.updated = updRows.length; r.removed = toRemove.length;
        log(`  ${r.added} ajout(s), ${r.updated} modification(s), ${r.removed} suppression(s)`);
        if (!dryRun) {
          progress({ step, total: ordered.length, tableId, phase: 'écriture' });
          if (updRows.length) await bulkWrite(docApi, 'BulkUpdateRecord', tableId, updIds, colIds, updRows, (d, t) => progress({ step, total: ordered.length, tableId, phase: 'modifications', done: d, of: t }));
          if (toAdd.length) await bulkWrite(docApi, 'BulkAddRecord', tableId, null, colIds, toAdd, (d, t) => progress({ step, total: ordered.length, tableId, phase: 'ajouts', done: d, of: t }));
          if (toRemove.length) await bulkRemove(docApi, tableId, toRemove);
          finalState[tableId] = await docApi.fetchTable(tableId);
        } else {
          // État simulé : identifiants existants + identifiants fictifs pour les ajouts
          const sim = { id: [] }; colIds.forEach((c) => { sim[c] = []; });
          const removed = new Set(toRemove);
          exIds.forEach((id, i) => { if (removed.has(id)) return; sim.id.push(id); colIds.forEach((c) => sim[c].push(existing[c] ? existing[c][i] : null)); });
          toAdd.forEach((row, n) => { sim.id.push(-(n + 1)); colIds.forEach((c, j) => sim[c].push(row[j])); });
          finalState[tableId] = sim;
        }
      } else {
        // Remplacement complet, seulement si le contenu a changé
        r.mode = 'remplacement complet';
        const sig = (rows) => rows.map((row) => JSON.stringify(row.map((v) => (v === '' || v === 0 ? null : v)))).sort().join('\n');
        const exRows = exIds.map((_, i) => colIds.map((c) => (existing[c] ? existing[c][i] : null)));
        if (exRows.length === newRows.length && sig(exRows) === sig(newRows)) {
          r.mode = 'inchangée'; r.skipped = true;
          log('  Contenu identique, aucune écriture.');
        } else {
          r.removed = exIds.length; r.added = newRows.length;
          log(`  ${r.removed} ligne(s) supprimée(s), ${r.added} ligne(s) ajoutée(s)`);
          if (!dryRun) {
            progress({ step, total: ordered.length, tableId, phase: 'suppression' });
            await bulkRemove(docApi, tableId, exIds, (d, t) => progress({ step, total: ordered.length, tableId, phase: 'suppression', done: d, of: t }));
            await bulkWrite(docApi, 'BulkAddRecord', tableId, null, colIds, newRows, (d, t) => progress({ step, total: ordered.length, tableId, phase: 'ajouts', done: d, of: t }));
            finalState[tableId] = await docApi.fetchTable(tableId);
          } else {
            const sim = { id: newRows.map((_, n) => -(n + 1)) };
            colIds.forEach((c, j) => { sim[c] = newRows.map((row) => row[j]); });
            finalState[tableId] = sim;
          }
        }
      }
      report.tables.push(r);
    }
    progress({ step: ordered.length, total: ordered.length, phase: 'terminé' });
    return report;
  }

  return { refresh, readSchema, readSheetMap, readColumnAliases, analyseWorkbook, orderPlans, convert, norm, keyOf, dateParts };
})();
if (typeof module !== 'undefined') module.exports = Refresher;
