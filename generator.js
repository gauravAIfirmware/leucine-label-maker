/*
 * Leucine label generator — a browser port of the backend's
 * utils/generateLeucineFromExcel.ts. Same template, same coordinates, same
 * font sizes. Works in the browser (window.LeucineLabels) and in Node
 * (module.exports) so it can be checked against the backend's output.
 *
 * Differences from the backend, all deliberate:
 *  - The sheet is chosen by the user, not hard-coded.
 *  - B CFU per strain comes from an explicit Adults (2) / Kids (1) choice.
 *  - The header row is checked against the 29 strain names; a shifted column
 *    would otherwise tick the wrong strains silently.
 *  - A text cell like "FALSE" or "No" counts as not selected (the backend
 *    treated any non-empty text as selected).
 *  - Name, batch and BS ID are trimmed.
 *  - The template and tick image are embedded once, not once per page.
 */
(function (root) {
  'use strict';

  const STRAINS = [
    'Lactobacillus lactis',
    'Bifidobacterium infantis',
    'Lactobacillus acidophilus',
    'Lactiplantibacillus plantarum',
    'Limosilactobacillus reuteri',
    'Lacticaseibacillus rhamnosus',
    'Ligilactobacillus salivarius',
    'Lacticaseibacillus casei',
    'Levilactobacillus brevis',
    'Lactobacillus johnsonii',
    'Bacillus coagulans',
    'Limosilactobacillus fermentum',
    'Lactobacillus caucasicus',
    'Lactobacillus helveticus',
    'Bifidobacterium bifidum',
    'Bifidobacterium breve',
    'Streptococcus thermophilus',
    'Saccharomyces boulardii',
    'Saccharomyces cerevisiae',
    'Lactobacillus gasseri',
    'Lacticaseibacillus paracasei',
    'Bacillus clausii',
    'Bacillus subtilis',
    'Bacillus indicus',
    'Lactobacillus delbrueckii',
    'Lactobacillus bulgaricus',
    'Bifidobacterium animalis',
    'Bifidobacterium lactis',
    'Bifidobacterium longum',
  ];

  const FIRST_STRAIN_COL = 4; // A=name, B=batch, C=MRP, D=BS ID, E.. = strains
  const DEFAULT_MRP = '12499';
  const SHELF_LIFE_MONTHS = 17;
  const CFU_PER_STRAIN = { adults: 2, kids: 1 };

  const TRUE_WORDS = new Set(['true', 'yes', 'y', '1', 'x', '✓', '✔', '✅']);
  const FALSE_WORDS = new Set(['false', 'no', 'n', '0', '-', '']);

  const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim().toLowerCase();

  /** @returns {{value: boolean, unknown: boolean}} */
  function readTick(v) {
    if (typeof v === 'boolean') return { value: v, unknown: false };
    if (typeof v === 'number') return { value: v !== 0, unknown: false };
    const s = norm(v);
    if (TRUE_WORDS.has(s)) return { value: true, unknown: false };
    if (FALSE_WORDS.has(s)) return { value: false, unknown: false };
    return { value: false, unknown: true };
  }

  function readWorkbook(XLSX, data) {
    return XLSX.read(data, { type: 'array' });
  }

  /** Best-guess audience from the sheet name, or null when it doesn't say. */
  function guessAudience(sheetName) {
    const n = norm(sheetName);
    if (/kid/.test(n)) return 'kids';
    if (/adult/.test(n)) return 'adults';
    return null;
  }

  /**
   * Read one sheet into patients. Errors block generation; warnings don't.
   */
  function parseSheet(XLSX, workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const errors = [];
    const warnings = [];
    if (!sheet) {
      return { patients: [], errors: [`Sheet "${sheetName}" was not found.`], warnings };
    }
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const header = rows[0] || [];

    const badCols = [];
    STRAINS.forEach((strain, i) => {
      const got = header[FIRST_STRAIN_COL + i];
      if (norm(got) !== norm(strain)) {
        badCols.push(`column ${XLSX.utils.encode_col(FIRST_STRAIN_COL + i)} should be "${strain}" but is "${String(got ?? '').trim() || '(empty)'}"`);
      }
    });
    if (badCols.length) {
      errors.push(
        `The strain columns don't line up with the label template, so ticks would land on the wrong strains. ` +
        `Expected columns A–D to be Name, Batch No, MRP, BS ID and E–AG to be the 29 strains in label order. ` +
        `First problem: ${badCols[0]}` + (badCols.length > 1 ? ` (and ${badCols.length - 1} more).` : '.')
      );
    }

    const patients = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.every((c) => String(c ?? '').trim() === '')) continue;
      const name = String(row[0] ?? '').trim();
      const batch = String(row[1] ?? '').trim();
      const price = String(row[2] ?? '').trim() || DEFAULT_MRP;
      const bsid = String(row[3] ?? '').trim();
      if (!name && !batch && !bsid) continue;

      const unknownCells = [];
      const selected = STRAINS.map((_, i) => {
        const t = readTick(row[FIRST_STRAIN_COL + i]);
        if (t.unknown) unknownCells.push(`${XLSX.utils.encode_col(FIRST_STRAIN_COL + i)}${r + 1} ("${String(row[FIRST_STRAIN_COL + i]).trim()}")`);
        return t.value;
      });
      if (unknownCells.length) {
        warnings.push(`Row ${r + 1} (${name || batch}): couldn't read ${unknownCells.join(', ')} as TRUE/FALSE — treated as not selected.`);
      }
      if (!batch) warnings.push(`Row ${r + 1} (${name || bsid}) has no batch number.`);
      if (!selected.some(Boolean)) warnings.push(`Row ${r + 1} (${name || batch}) has no strains selected.`);

      patients.push({ rowNumber: r + 1, name, batch, price, bsid, selected, strainCount: selected.filter(Boolean).length });
    }
    if (!patients.length && !errors.length) errors.push(`Sheet "${sheetName}" has no patient rows.`);
    return { patients, errors, warnings };
  }

  const formatMonth = (d) => d.toLocaleString('en-US', { month: 'short', year: 'numeric' });

  /** mfg is a Date; uses the 1st of its month so month arithmetic never overflows. */
  function labelDates(mfg) {
    const m = new Date(mfg.getFullYear(), mfg.getMonth(), 1);
    const e = new Date(m.getFullYear(), m.getMonth() + SHELF_LIFE_MONTHS, 1);
    return { mfg: formatMonth(m), exp: formatMonth(e) };
  }

  /**
   * Build the combined PDF. Returns Uint8Array.
   * opts: { patients, audience: 'adults'|'kids', mfgDate: Date, templateBytes, checkBytes }
   */
  async function buildPdf(PDFLib, opts) {
    const { PDFDocument, StandardFonts, rgb } = PDFLib;
    const rate = CFU_PER_STRAIN[opts.audience];
    if (!rate) throw new Error('Choose Adults or Kids first.');

    const pdfDoc = await PDFDocument.create();
    const template = await PDFDocument.load(opts.templateBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const check = await pdfDoc.embedPng(opts.checkBytes);
    const checkH = 6;
    const checkW = checkH * (check.width / check.height);
    const black = rgb(0, 0, 0);
    const { mfg, exp } = labelDates(opts.mfgDate);

    for (const p of opts.patients) {
      const [tpl] = await pdfDoc.copyPages(template, [0]);
      const page = pdfDoc.addPage(tpl);
      const text = (x, y, size, s) => page.drawText(s, { x, y, size, font, color: black });

      text(376, 707, 10, `Batch No: ${p.batch}`);
      text(376, 695, 10, `Mfg Date: ${mfg}`);
      text(376, 683, 10, `Exp Date: ${exp}`);
      text(376, 671, 10, `MRP:        ${p.price}`);
      text(188, 690, 5, `${p.strainCount * rate} B CFU`);
      text(110, 520, 10, p.name);
      text(110, 502, 12, p.bsid);

      let y = 410;
      p.selected.forEach((on, i) => {
        if (on) page.drawImage(check, { x: 242, y, width: checkW, height: checkH });
        y -= 10;
        if (i === 12) y -= 2;
      });

      y = 412;
      p.selected.forEach((on, i) => {
        text(239, y, 5, `[   ${on ? '' : '-'}   ]`);
        y -= 10;
        if (i === 12) y -= 2;
      });
    }
    return pdfDoc.save();
  }

  const api = { STRAINS, CFU_PER_STRAIN, readWorkbook, guessAudience, parseSheet, labelDates, buildPdf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LeucineLabels = api;
})(typeof window !== 'undefined' ? window : globalThis);
