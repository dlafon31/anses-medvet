/* Lecture du classeur Excel hors du fil principal : la page reste réactive
 * (préchargeur animé) pendant l'analyse d'un fichier de plusieurs dizaines de Mo. */
importScripts('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');

self.onmessage = (e) => {
  try {
    self.postMessage({ type: 'step', text: 'Décodage du fichier Excel…' });
    const wb = XLSX.read(new Uint8Array(e.data), { type: 'array', cellDates: false });
    const aoa = {};
    wb.SheetNames.forEach((name, i) => {
      self.postMessage({ type: 'step', text: `Lecture de l'onglet « ${name} » (${i + 1}/${wb.SheetNames.length})…`, ratio: i / wb.SheetNames.length });
      aoa[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
    });
    self.postMessage({ type: 'done', SheetNames: wb.SheetNames, aoa });
  } catch (err) {
    self.postMessage({ type: 'error', message: String((err && err.message) || err) });
  }
};
