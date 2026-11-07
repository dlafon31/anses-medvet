/* ==========================================================================
 * Widget Grist — actualisation de la base des médicaments vétérinaires ANSES
 * Interface : téléchargement de la copie du fichier, analyse, application.
 * Dépend de grist-plugin-api.js, de SheetJS (XLSX) et de refresher.js.
 * ========================================================================== */
(() => {
  'use strict';
  // Copie publiée chaque nuit par GitHub Actions (le serveur de l'ANSES n'autorise pas les requêtes CORS)
  const DATAGOUV_URL = 'https://www.data.gouv.fr/api/1/datasets/r/7a6e415d-1627-4051-963f-4d572ec27e5c';
  const DEFAULT_URL = 'https://dlafon31.github.io/anses-medvet/data/amm-vet-fr.xls';
  const $ = (id) => document.getElementById(id);
  const norm = Refresher.norm;
  let workbook = null; let sourceLabel = ''; let busy = false;

  const fmtDate = (d) => d ? new Date(d).toLocaleString('fr-FR', { dateStyle: 'long', timeStyle: 'short' }) : '—';
  const fmtSize = (b) => b ? (b / 1048576).toFixed(1).replace('.', ',') + ' Mo' : '';

  function log(msg, cls) {
    const line = document.createElement('div');
    if (cls) line.className = cls;
    line.textContent = `[${new Date().toLocaleTimeString('fr-FR')}] ${msg}`;
    $('log').appendChild(line); $('log').scrollTop = $('log').scrollHeight;
  }
  function banner(msg, kind) {
    const b = $('banner');
    if (!msg) { b.className = 'banner hidden'; return; }
    b.className = 'banner ' + kind; b.innerHTML = msg;
  }
  function setProgress(ratio, text) {
    $('bar').style.width = Math.round(Math.max(0, Math.min(1, ratio)) * 100) + '%';
    if (text !== undefined) $('phase').textContent = text;
  }
  function setBusy(b) {
    busy = b;
    $('btnAnalyse').disabled = b; $('file').disabled = b;
    $('btnApply').disabled = b || !workbook;
  }

  // ---------- Paramètres ----------
  grist.ready({ requiredAccess: 'full', allowSelectBy: false });
  grist.onOptions((opts) => { if (opts && opts.sourceUrl && !/data\.gouv\.fr|anses\.fr/.test(opts.sourceUrl)) $('url').value = opts.sourceUrl; refreshInfo(); });
  $('url').value = DEFAULT_URL;
  refreshInfo();
  $('url').addEventListener('change', async () => {
    try { await grist.setOption('sourceUrl', $('url').value.trim() || DEFAULT_URL); } catch (e) { /* lecture seule */ }
    refreshInfo();
  });

  // ---------- Informations ----------
  async function refreshInfo() {
    const url = $('url').value.trim();
    const rid = (url.match(/\/r\/([0-9a-f-]{36})/i) || [])[1];
    let metaOk = false;
    if (!rid) { // miroir : métadonnées dans meta.json à côté du fichier
      try {
        const r = await fetch(new URL('meta.json', url), { cache: 'no-store' });
        const m = await r.json();
        $('remote').textContent = `copie du ${fmtDate(m.fetched_at)}${m.source_last_modified ? ' · source modifiée le ' + fmtDate(m.source_last_modified) : ''}${m.size ? ' · ' + fmtSize(m.size) : ''}`;
        metaOk = true;
      } catch (e) { /* pas de meta.json */ }
    }
    if (metaOk) { /* déjà affiché */ } else if (rid) {
      try {
        const r = await fetch(`https://www.data.gouv.fr/api/2/datasets/resources/${rid}/`);
        const j = await r.json();
        const res = j.resource || j;
        $('remote').textContent = `${fmtDate(res.last_modified)} ${res.filesize ? '· ' + fmtSize(res.filesize) : ''}`;
      } catch (e) { $('remote').textContent = 'indisponible'; }
    } else $('remote').textContent = '—';
    try {
      const t = await grist.docApi.fetchTable('Date_Export_BDD');
      const v = (t.Date_Export || []).find((x) => typeof x === 'number');
      $('local').textContent = v ? fmtDate(v * 1000) : '—';
    } catch (e) { $('local').textContent = '—'; }
  }

  // ---------- Chargement du classeur ----------
  async function download(url) {
    setProgress(0, 'Téléchargement du fichier…');
    const resp = await fetch(url, { redirect: 'follow', cache: 'no-store' });
    if (!resp.ok) throw new Error(`HTTP ${resp.status} sur ${url}`);
    const total = +resp.headers.get('content-length') || 0;
    if (!resp.body || !total) return new Uint8Array(await resp.arrayBuffer());
    const reader = resp.body.getReader(); const chunks = []; let got = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value); got += value.length;
      setProgress(got / total, `Téléchargement… ${fmtSize(got)} / ${fmtSize(total)}`);
    }
    const out = new Uint8Array(got); let o = 0;
    for (const c of chunks) { out.set(c, o); o += c.length; }
    return out;
  }

  function parse(bytes) {
    setProgress(1, 'Lecture du classeur Excel…');
    return XLSX.read(bytes, { type: 'array', cellDates: false });
  }

  // ---------- Exécution ----------
  function renderReport(rep) {
    const tb = $('result'); tb.innerHTML = '';
    for (const t of rep.tables) {
      const alerts = [];
      if (t.badRefs) alerts.push(`${t.badRefs} réf. orphelines`);
      if (t.badDates) alerts.push(`${t.badDates} dates illisibles`);
      if (t.unmatchedHeaders.length) alerts.push(`colonnes Excel ignorées : ${t.unmatchedHeaders.join(', ')}`);
      if (t.unmappedCols.length) alerts.push(`absentes de l'Excel : ${t.unmappedCols.join(', ')}`);
      const changed = t.added || t.updated || t.removed;
      const tr = document.createElement('tr');
      tr.innerHTML = `<td><b></b></td><td><span class="tag ${t.blocked ? 'err' : changed ? 'chg' : 'ok'}"></span></td>
        <td class="n">${t.rowsExcel.toLocaleString('fr-FR')}</td><td class="n">${t.added.toLocaleString('fr-FR')}</td>
        <td class="n">${t.updated.toLocaleString('fr-FR')}</td><td class="n">${t.removed.toLocaleString('fr-FR')}</td><td class="warn"></td>`;
      tr.querySelector('b').textContent = t.tableId;
      if (norm(t.sheet) !== norm(t.tableId)) tr.firstElementChild.append(document.createTextNode(` ← « ${t.sheet} »`));
      tr.querySelector('.tag').textContent = t.mode;
      tr.lastElementChild.textContent = alerts.join(' · ');
      tb.appendChild(tr);
    }
    if (rep.untouchedTables.length) {
      const tr = document.createElement('tr');
      tr.innerHTML = '<td colspan="7" class="k"></td>';
      tr.firstElementChild.textContent = `Tables non concernées (aucun onglet correspondant) : ${rep.untouchedTables.join(', ')}`;
      tb.appendChild(tr);
    }
    $('resultCard').classList.remove('hidden');
  }

  async function run(dryRun) {
    const rep = await Refresher.refresh({
      docApi: grist.docApi, XLSX, workbook, dryRun,
      log: (m, cls) => log(m, cls),
      progress: (p) => {
        const base = (p.step - 1) / p.total;
        const inner = p.of ? (p.done / p.of) / p.total : 0;
        setProgress(p.phase === 'terminé' ? 1 : base + inner,
          p.phase === 'terminé' ? '' : `${p.tableId} — ${p.phase}${p.of ? ` ${p.done.toLocaleString('fr-FR')} / ${p.of.toLocaleString('fr-FR')}` : ''} (${p.step}/${p.total})`);
      },
    });
    renderReport(rep);
    return rep;
  }

  async function analyse(loader) {
    if (busy) return;
    setBusy(true); banner(''); workbook = null;
    try {
      const { bytes, label } = await loader();
      workbook = parse(bytes); sourceLabel = label;
      log(`Classeur chargé (${label}) : onglets ${workbook.SheetNames.join(', ')}`);
      log('Analyse des différences (aucune écriture)…');
      const rep = await run(true);
      const nChanges = rep.tables.reduce((s, t) => s + t.added + t.updated + t.removed, 0);
      setProgress(1, nChanges ? `Analyse terminée : ${nChanges.toLocaleString('fr-FR')} opération(s) à appliquer.` : 'Analyse terminée : le document est déjà à jour.');
      banner(nChanges ? 'Vérifiez le récapitulatif ci-dessous puis cliquez sur <b>2. Appliquer les modifications</b>.' : 'Le document est déjà à jour, rien à appliquer.', 'info');
      if (!nChanges) workbook = null;
    } catch (e) {
      workbook = null;
      log('Erreur : ' + (e.message || e), 'error');
      const cors = e instanceof TypeError; // « Failed to fetch » / « NetworkError » = blocage CORS ou réseau
      banner(`Le téléchargement automatique a échoué (${cors ? 'accès refusé par le serveur source (CORS) ou réseau indisponible' : (e.message || e)}).<br>
        Téléchargez le fichier depuis <a href="${DATAGOUV_URL}" target="_blank" rel="noopener">data.gouv.fr</a>,
        puis utilisez <b>Importer un fichier Excel…</b>`, 'error');
      setProgress(0, '');
    } finally { setBusy(false); }
  }

  $('btnAnalyse').addEventListener('click', () => analyse(async () => {
    const url = $('url').value.trim() || DEFAULT_URL;
    return { bytes: await download(url), label: 'téléchargé depuis data.gouv.fr' };
  }));

  $('file').addEventListener('change', (ev) => {
    const f = ev.target.files[0]; ev.target.value = '';
    if (f) analyse(async () => ({ bytes: new Uint8Array(await f.arrayBuffer()), label: f.name }));
  });

  $('btnApply').addEventListener('click', async () => {
    if (busy || !workbook) return;
    setBusy(true); banner('');
    const t0 = Date.now();
    try {
      log(`Application des modifications (${sourceLabel})…`);
      const rep = await run(false);
      const s = Math.round((Date.now() - t0) / 1000);
      log(`Actualisation terminée en ${s} s.`, 'okc');
      banner(`Actualisation terminée en ${s} s.`, 'info');
      setProgress(1, 'Terminé.');
    } catch (e) {
      log('Erreur pendant l\'écriture : ' + (e.message || e), 'error');
      banner('Erreur pendant l\'écriture : ' + (e.message || e) + '<br>Les lots déjà envoyés sont enregistrés ; relancez l\'actualisation pour terminer (elle reprend là où les données diffèrent), ou restaurez une version via l\'historique du document.', 'error');
    } finally { workbook = null; setBusy(false); refreshInfo(); }
  });
})();
