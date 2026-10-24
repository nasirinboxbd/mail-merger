/* Core mail-merge logic (no DOM). Works in the browser (globals PizZip, XLSX) and in Node (for tests). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('pizzip'), require('xlsx'));
  else root.MailMerge = factory(root.PizZip, root.XLSX);
})(typeof self !== 'undefined' ? self : this, function (PizZip, XLSX) {
  'use strict';
  const MAX_RECORDS = 2000;
  const userErr = (m) => Object.assign(new Error(m), { user: true });

  const PARA_RE = /<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g; // innermost paragraphs
  const RUN_RE = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
  const PH_RE = /<<(\s*)([^<>\t\n\u0000]*?\S)(\s*)>>/g;
  const TOKEN_SRC = '<w:t(?:\\s[^>]*)?>([\\s\\S]*?)<\\/w:t>|<w:t\\s*\\/>|(<w:tab\\s*\\/>)|(<w:br(?:\\s[^>]*)?\\/>)|<(\\w+:\\w+)(?:\\s[^>]*)?\\/>|<(\\w+:\\w+)(?:\\s[^>]*)?>[\\s\\S]*?<\\/\\5>';

  const normalizeKey = (s) => String(s).toLowerCase().trim().replace(/[\s\-_.\/\\]+/g, '');
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const dec = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi, (m, e) => {
    const l = e.toLowerCase();
    if (l === 'lt') return '<'; if (l === 'gt') return '>'; if (l === 'amp') return '&';
    if (l === 'quot') return '"'; if (l === 'apos') return "'";
    return String.fromCodePoint(l[1] === 'x' ? parseInt(l.slice(2), 16) : parseInt(l.slice(1), 10));
  });

  function parseRun(xml) {
    const open = /^<w:r(?:\s[^>]*)?>/.exec(xml)[0];
    let rest = xml.slice(open.length, xml.length - 6);
    const rm = /^\s*(<w:rPr(?:\s[^>]*)?\/>|<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>)/.exec(rest);
    const rPr = rm ? rm[1] : '';
    if (rm) rest = rest.slice(rm[0].length);
    const units = [];
    if (rest.replace(new RegExp(TOKEN_SRC, 'g'), '').trim()) {
      units.push({ ch: '\u0000', raw: rest }); // unknown content: keep untouched
    } else {
      for (const m of rest.matchAll(new RegExp(TOKEN_SRC, 'g'))) {
        if (m[1] !== undefined) for (const ch of dec(m[1])) units.push({ ch });
        else if (m[2]) units.push({ ch: '\t', raw: m[0] });
        else if (m[3]) units.push({ ch: '\n', raw: m[0] });
        else if (m[4] || m[5]) units.push({ ch: '\u0000', raw: m[0] });
      }
    }
    return { open, rPr, units };
  }

  function serializeRun(r) {
    let out = r.open + r.rPr, buf = '';
    const flush = () => { if (buf) { out += '<w:t xml:space="preserve">' + esc(buf) + '</w:t>'; buf = ''; } };
    for (const u of r.units) { if (u.raw) { flush(); out += u.raw; } else buf += u.ch; }
    flush();
    return out + '</w:r>';
  }

  function toUnits(value) {
    const units = [];
    for (const ch of String(value).replace(/\r\n?/g, '\n')) {
      if (ch === '\n') units.push({ ch, raw: '<w:br/>' });
      else if (ch === '\t') units.push({ ch, raw: '<w:tab/>' });
      else units.push({ ch });
    }
    return units;
  }

  function processParagraph(p, ctx) {
    const ms = [...p.matchAll(RUN_RE)];
    if (!ms.length) return p;
    const runs = ms.map((m) => ({ start: m.index, end: m.index + m[0].length, ...parseRun(m[0]) }));
    let stream = '';
    for (const r of runs) for (const u of r.units) stream += u.ch;
    if (!stream.includes('<<')) return p;
    const del = new Set(), ins = new Map();
    for (const m of stream.matchAll(PH_RE)) {
      const key = normalizeKey(m[2]);
      ctx.found(m[2], key);
      if (ctx.values) {
        for (let i = m.index; i < m.index + m[0].length; i++) del.add(i);
        // the value goes where the field NAME starts, so it inherits the name run's formatting
        ins.set(m.index + 2 + m[1].length, toUnits(ctx.values[key] || ''));
      }
    }
    if (!del.size) return p;
    let out = '', pos = 0, g = 0;
    for (const r of runs) {
      const nu = []; let changed = false;
      for (const u of r.units) {
        if (ins.has(g)) { nu.push(...ins.get(g)); changed = true; }
        if (del.has(g)) changed = true; else nu.push(u);
        g++;
      }
      out += p.slice(pos, r.start) + (changed ? serializeRun({ ...r, units: nu }) : p.slice(r.start, r.end));
      pos = r.end;
    }
    return out + p.slice(pos);
  }

  const processXml = (xml, ctx) => xml.replace(PARA_RE, (p) => processParagraph(p, ctx));
  const mergeParts = (zip) => Object.keys(zip.files).filter((n) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n));

  function openZip(buf) {
    let zip;
    try { zip = new PizZip(buf); } catch (e) { throw userErr('The Word file is not a valid .docx (it cannot be opened as a ZIP package).'); }
    if (!zip.file('word/document.xml')) throw userErr('The Word file has no word/document.xml. Is it a real .docx (not .doc or a renamed file)?');
    return zip;
  }

  function detectPlaceholders(buf) {
    const zip = openZip(buf), map = new Map();
    for (const part of mergeParts(zip)) {
      const label = part === 'word/document.xml' ? 'body' : part.replace('word/', '').replace('.xml', '');
      processXml(zip.file(part).asText(), { found: (name, key) => {
        const e = map.get(key) || { name: name.trim(), key, count: 0, locations: new Set() };
        e.count++; e.locations.add(label); map.set(key, e);
      } });
    }
    return [...map.values()].map((e) => ({ ...e, locations: [...e.locations] }));
  }

  function splitBody(xml) {
    const a = xml.indexOf('<w:body>'), b = xml.lastIndexOf('</w:body>');
    if (a < 0 || b < 0) throw userErr('Unsupported document structure (no w:body).');
    let inner = xml.slice(a + 8, b), sect = '';
    const i = inner.lastIndexOf('<w:sectPr');
    if (i >= 0 && !/<\/w:p>|<\/w:tbl>/.test(inner.slice(i))) { sect = inner.slice(i); inner = inner.slice(0, i); }
    return { head: xml.slice(0, a + 8), inner, sect, tail: xml.slice(b) };
  }

  /** records: array of { normalizedKey: value }. Returns a Uint8Array (.docx) with one page-separated copy per record. */
  function mergeDocx(buf, records) {
    const zip = openZip(buf);
    const first = records[0] || {};
    for (const part of mergeParts(zip)) {
      const xml = zip.file(part).asText();
      if (part === 'word/document.xml') {
        const { head, inner, sect, tail } = splitBody(xml);
        const brk = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
        const body = (records.length ? records : [{}]).map((rec) => processXml(inner, { found() {}, values: rec })).join(brk);
        zip.file(part, head + body + sect + tail);
      } else {
        zip.file(part, processXml(xml, { found() {}, values: first })); // headers/footers are shared: first record
      }
    }
    return zip.generate({ type: 'uint8array', compression: 'DEFLATE' });
  }

  /** Preview-only copy: turns run colours and paragraph alignment into styles Mammoth can map to classes. */
  function prepareForPreview(buf) {
    const zip = openZip(buf);
    let xml = zip.file('word/document.xml').asText();
    const colors = new Set(), aligns = new Set();
    xml = xml.replace(RUN_RE, (r) => {
      const m = /<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>/.exec(r);
      if (!m || m[0].includes('<w:rStyle')) return r;
      const c = /<w:color\s+w:val="([0-9A-Fa-f]{6})"/.exec(m[0]);
      if (!c) return r;
      const hex = c[1].toUpperCase(); colors.add(hex);
      return r.replace(/<w:rPr(\s[^>]*)?>/, (o) => o + `<w:rStyle w:val="Clr_${hex}"/>`);
    });
    xml = xml.replace(PARA_RE, (p) => {
      const m = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p);
      if (!m || m[0].includes('<w:pStyle')) return p;
      const j = /<w:jc\s+w:val="(center|right|both|end)"/.exec(m[0]);
      if (!j) return p;
      const a = j[1] === 'end' ? 'right' : j[1]; aligns.add(a);
      return p.replace('<w:pPr>', `<w:pPr><w:pStyle w:val="Jc_${a}"/>`);
    });
    zip.file('word/document.xml', xml);
    const styleMap = ['u => u'];
    const sf = zip.file('word/styles.xml');
    if (sf) {
      let add = '';
      for (const h of colors) { add += `<w:style w:type="character" w:styleId="Clr_${h}"><w:name w:val="Clr ${h}"/></w:style>`; styleMap.push(`r[style-name='Clr ${h}'] => span.clr-${h}`); }
      for (const a of aligns) { add += `<w:style w:type="paragraph" w:styleId="Jc_${a}"><w:name w:val="Jc ${a}"/></w:style>`; styleMap.push(`p[style-name='Jc ${a}'] => p.jc-${a}:fresh`); }
      zip.file('word/styles.xml', sf.asText().replace('</w:styles>', add + '</w:styles>'));
    }
    return { arrayBuffer: zip.generate({ type: 'arraybuffer' }), styleMap };
  }

  function readWorkbook(buf) {
    let wb;
    try { wb = XLSX.read(new Uint8Array(buf), { type: 'array' }); } catch (e) { throw userErr('The Excel file could not be read. It may be corrupt or password-protected.'); }
    const sheetName = wb.SheetNames[0];
    if (!sheetName) throw userErr('The Excel workbook contains no worksheets.');
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: false, defval: '', blankrows: false });
    if (!rows.length) throw userErr(`The first worksheet ("${sheetName}") is empty. Put column headers in row 1.`);
    const warnings = [], seen = new Map(), columns = [], keyToCol = new Map();
    rows[0].forEach((h, index) => {
      let header = String(h).trim();
      if (!header) return;
      const key = normalizeKey(header);
      if (seen.has(key)) {
        seen.set(key, seen.get(key) + 1);
        warnings.push(`Duplicate column "${header}" — only the first one is used for matching.`);
        header += ` (${seen.get(key)})`;
      } else { seen.set(key, 1); keyToCol.set(key, index); }
      columns.push({ index, header, key });
    });
    if (!columns.length) throw userErr('Row 1 of the first worksheet has no headers.');
    const data = rows.slice(1);
    if (data.length > MAX_RECORDS) throw userErr(`The workbook has ${data.length} records; the limit is ${MAX_RECORDS}.`);
    const records = data.map((r) => { const o = {}; for (const [k, i] of keyToCol) o[k] = String(r[i] ?? '').trim(); return o; });
    const rowLabels = data.map((r) => String(r.find((c) => String(c).trim() !== '') ?? '').trim().slice(0, 40));
    if (!records.length) warnings.push('The worksheet has headers but no data rows.');
    return { sheetName, columns, records, rowLabels, warnings };
  }

  /** Which placeholders map to which Excel columns. */
  function analyze(placeholders, wb) {
    const out = { mapping: [], missingColumns: [], unusedColumns: [] };
    if (!wb) return out;
    const byKey = new Map(wb.columns.map((c) => [c.key, c.header]));
    out.mapping = placeholders.map((p) => ({ name: p.name, key: p.key, column: byKey.get(p.key) || null }));
    out.missingColumns = out.mapping.filter((m) => !m.column).map((m) => m.name);
    const used = new Set(placeholders.map((p) => p.key));
    out.unusedColumns = wb.columns.filter((c) => !used.has(c.key)).map((c) => c.header);
    return out;
  }

  const base = (keys, row) => Object.fromEntries(keys.map((k) => [k, (row && row[k]) || '']));

  /** ov: { key: string } website overrides. Used by the final download. */
  function buildRecords(rows, keys, ov, mode) {
    if (mode === 'append') return [...rows.map((r) => base(keys, r)), { ...base(keys, null), ...ov }];
    return rows.length ? rows.map((r) => ({ ...base(keys, r), ...ov })) : [{ ...base(keys, null), ...ov }];
  }

  /** The records shown in the live preview (selected row + optional extra record). Same rules as buildRecords. */
  function previewRecords(rows, keys, ov, mode, idx) {
    let rowRec = null;
    const i = Math.min(Math.max(idx || 0, 0), Math.max(rows.length - 1, 0));
    if (rows.length) rowRec = mode === 'replace' ? { ...base(keys, rows[i]), ...ov } : base(keys, rows[i]);
    else if (mode === 'replace') rowRec = { ...base(keys, null), ...ov };
    return { rowRec, extraRec: mode === 'append' ? { ...base(keys, null), ...ov } : null, rowNumber: rows.length ? i + 1 : null };
  }

  return { normalizeKey, detectPlaceholders, mergeDocx, prepareForPreview, readWorkbook, analyze, buildRecords, previewRecords, MAX_RECORDS };
});
