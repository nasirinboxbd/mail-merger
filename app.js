(function () {
  'use strict';
  const MM = window.MailMerge, MAX_MB = 10;
  const $ = (s, r = document) => r.querySelector(s);
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const state = { tplBuf: null, tplName: '', placeholders: [], wbBuf: null, wbName: '', wb: null, analysis: null,
    values: {}, mode: 'replace', row: 0, pvTimer: null, pvToken: 0 };

  /* ---------- helpers ---------- */
  function message(kind, text) { const d = document.createElement('div'); d.className = 'msg ' + kind; d.textContent = text; $('#messages').appendChild(d); }
  function clearMessages() { $('#messages').innerHTML = ''; }
  function busy(on, text) { $('#busy').hidden = !on; if (text) $('#busyText').textContent = text; }
  const friendly = (e) => (e && e.user ? e.message : 'Something went wrong while processing the file. Please check that it is valid and try again.');

  function overrides() {
    const o = {};
    for (const [k, v] of Object.entries(state.values)) {
      if (v.cleared) o[k] = ''; else if (v.value !== '') o[k] = v.value;
    }
    return o;
  }
  const keys = () => state.placeholders.map((p) => p.key);
  function updateGenerate() { $('#generate').disabled = !(state.tplBuf && (state.wb || Object.keys(overrides()).length)); }

  /* ---------- file upload ---------- */
  function validate(file, exts, kind) {
    const ext = (file.name.match(/\.[^.]+$/) || [''])[0].toLowerCase();
    if (!exts.includes(ext)) throw Object.assign(new Error(`"${file.name}": expected ${exts.join(', ')} for the ${kind}.`), { user: true });
    if (file.size > MAX_MB * 1024 * 1024) throw Object.assign(new Error(`"${file.name}" is too large (limit ${MAX_MB} MB).`), { user: true });
  }
  function magicOk(buf, kind) {
    const b = new Uint8Array(buf); if (b.length < 8) return false;
    const zip = b[0] === 0x50 && b[1] === 0x4b, ole = b[0] === 0xd0 && b[1] === 0xcf;
    return kind === 'template' ? zip : zip || ole;
  }

  function wireDrop(zoneId, exts, kind, onBuffer, onRemove) {
    const zone = $(zoneId), input = $('input[type=file]', zone);
    const show = (name) => { $('.drop-empty', zone).hidden = !!name; $('.file-row', zone).hidden = !name; $('.file-name', zone).textContent = name ? '✔ ' + name : ''; zone.classList.toggle('has', !!name); };
    async function take(file) {
      if (!file) return;
      clearMessages();
      try {
        validate(file, exts, kind);
        const buf = await file.arrayBuffer();
        if (!magicOk(buf, kind)) throw Object.assign(new Error(kind === 'template'
          ? 'The Word file does not look like a valid .docx (legacy .doc files are not supported — save as .docx).'
          : 'The Excel file does not look like a valid workbook.'), { user: true });
        busy(true, 'Inspecting ' + kind + '…');
        await new Promise((r) => setTimeout(r, 20)); // let the spinner paint
        onBuffer(buf, file.name);
        show(file.name);
      } catch (e) { message('error', friendly(e)); if (!e.user) console.error(e); }
      busy(false);
      afterInspect();
    }
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('over'));
    zone.addEventListener('drop', (e) => { e.preventDefault(); zone.classList.remove('over'); take(e.dataTransfer.files[0]); });
    zone.querySelectorAll('[data-pick]').forEach((b) => b.addEventListener('click', () => input.click()));
    input.addEventListener('change', () => { take(input.files[0]); input.value = ''; });
    $('[data-remove]', zone).addEventListener('click', () => { clearMessages(); onRemove(); show(''); afterInspect(); });
    return { show };
  }

  const tplZone = wireDrop('#dropTpl', ['.docx'], 'template', (buf, name) => {
    const placeholders = MM.detectPlaceholders(new Uint8Array(buf)); // throws on bad file before state changes
    state.tplBuf = new Uint8Array(buf); state.tplName = name; state.placeholders = placeholders; state.row = 0;
    if (!placeholders.length) message('warn', 'No <<placeholders>> were found in the template body, headers or footers.');
  }, () => { state.tplBuf = null; state.placeholders = []; });
  const wbZone = wireDrop('#dropWb', ['.xlsx', '.xls', '.xlsm'], 'workbook', (buf, name) => {
    const wb = MM.readWorkbook(buf);
    state.wbBuf = buf; state.wbName = name; state.wb = wb; state.row = 0;
    wb.warnings.forEach((w) => message('warn', w));
  }, () => { state.wb = null; state.wbBuf = null; });

  /* ---------- rendering ---------- */
  function afterInspect() {
    state.analysis = MM.analyze(state.placeholders, state.wb);
    // keep only values for placeholders that still exist
    for (const k of Object.keys(state.values)) if (!keys().includes(k)) delete state.values[k];
    renderExcel(); renderFields(); renderRowSelect(); updateGenerate(); schedulePreview();
  }

  function renderExcel() {
    const c = $('#excelCard'), wb = state.wb;
    c.hidden = !wb; if (!wb) return;
    const a = state.analysis;
    c.innerHTML = `<h2>Excel data</h2><p>Sheet “${esc(wb.sheetName)}” · <b>${wb.records.length}</b> record(s)</p>
      <p>${wb.columns.map((x) => `<span class="chip">${esc(x.header)}</span>`).join('')}</p>
      ${a.missingColumns.length ? `<p class="msg warn">No Excel column for: ${a.missingColumns.map((m) => esc('<<' + m + '>>')).join(', ')}. These stay blank unless you type a value.</p>` : ''}
      ${a.unusedColumns.length ? `<p class="muted">Columns not used by the template: ${a.unusedColumns.map(esc).join(', ')}</p>` : ''}`;
  }

  function renderFields() {
    const card = $('#phCard'), box = $('#fields');
    card.hidden = !state.tplBuf; box.innerHTML = '';
    if (!state.tplBuf) return;
    $('#phTitle').textContent = `Placeholders (${state.placeholders.length})`;
    if (!state.placeholders.length) box.innerHTML = '<p class="muted">None detected.</p>';
    state.placeholders.forEach((p) => {
      const m = state.analysis.mapping.find((x) => x.key === p.key), hasCol = !!(m && m.column);
      const v = state.values[p.key] || (state.values[p.key] = { value: '', cleared: false });
      const wrap = document.createElement('div'); wrap.className = 'field';
      wrap.innerHTML = `<div class="field-head"><label for="f-${esc(p.key)}">${esc(p.name)}</label>
        ${state.wb ? (hasCol ? `<span class="badge ok">Excel: ${esc(m.column)}</span>` : '<span class="badge bad">No Excel column</span>') : ''}
        <span class="badge">${esc(p.locations.join(', '))} ×${p.count}</span></div>
        <div class="field-row"><input id="f-${esc(p.key)}" type="text" autocomplete="off" /><button type="button" class="small"></button></div>`;
      const input = $('input', wrap), btn = $('button', wrap);
      const sync = () => {
        input.disabled = v.cleared; input.value = v.cleared ? '' : v.value;
        input.placeholder = v.cleared ? '(intentionally blank)' : hasCol ? 'Empty = use Excel value' : 'Type a value';
        btn.textContent = v.cleared ? 'Undo clear' : 'Set blank';
        btn.title = v.cleared ? 'Stop using an intentionally blank value' : 'Use an intentionally empty value';
      };
      input.addEventListener('input', () => { v.value = input.value; v.cleared = false; updateGenerate(); schedulePreview(); });
      btn.addEventListener('click', () => { v.cleared = !v.cleared; v.value = ''; sync(); updateGenerate(); schedulePreview(); });
      sync(); box.appendChild(wrap);
    });
  }

  function renderRowSelect() {
    const wrap = $('#rowWrap'), sel = $('#rowSel'), wb = state.wb;
    wrap.hidden = !(wb && wb.records.length); sel.innerHTML = '';
    if (wrap.hidden) return;
    wb.rowLabels.forEach((l, i) => { const o = document.createElement('option'); o.value = i; o.textContent = (i + 1) + (l ? ' — ' + l : ''); sel.appendChild(o); });
    sel.value = state.row;
  }
  $('#rowSel').addEventListener('change', (e) => { state.row = +e.target.value; schedulePreview(); });
  document.querySelectorAll('input[name=mode]').forEach((r) => r.addEventListener('change', () => { state.mode = r.value; schedulePreview(); }));

  /* ---------- preview ---------- */
  function cleanHtml(html) {
    // Turn Mammoth's helper classes into validated inline styles, drop everything else, then sanitise.
    const doc = new DOMParser().parseFromString('<body>' + html + '</body>', 'text/html');
    doc.body.querySelectorAll('*').forEach((el) => {
      const cls = el.getAttribute('class') || '', style = [];
      const c = /clr-([0-9A-F]{6})/.exec(cls); if (c) style.push('color:#' + c[1]);
      const j = /jc-(center|right|both)/.exec(cls); if (j) style.push('text-align:' + (j[1] === 'both' ? 'justify' : j[1]));
      for (const a of [...el.attributes]) if (!['href', 'src', 'alt', 'colspan', 'rowspan'].includes(a.name)) el.removeAttribute(a.name);
      if (style.length) el.setAttribute('style', style.join(';'));
    });
    return DOMPurify.sanitize(doc.body.innerHTML, {
      ALLOWED_TAGS: ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'strong', 'em', 'u', 's', 'sub', 'sup', 'br', 'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tr', 'td', 'th', 'span', 'a', 'img'],
      ALLOWED_ATTR: ['style', 'href', 'src', 'alt', 'colspan', 'rowspan'],
    });
  }
  async function toHtml(record) {
    const merged = MM.mergeDocx(state.tplBuf, [record]);
    const { arrayBuffer, styleMap } = MM.prepareForPreview(merged);
    const res = await mammoth.convertToHtml({ arrayBuffer }, { styleMap });
    return cleanHtml(res.value);
  }
  function schedulePreview() { clearTimeout(state.pvTimer); state.pvTimer = setTimeout(renderPreview, 250); }

  async function renderPreview() {
    const token = ++state.pvToken, rowBox = $('#pvRow'), extraBox = $('#pvExtra'), err = $('#pvError');
    $('#pvEmpty').hidden = !!state.tplBuf;
    if (!state.tplBuf) { rowBox.innerHTML = extraBox.innerHTML = err.innerHTML = ''; return; }
    $('#pvSpin').hidden = false;
    try {
      const rows = state.wb ? state.wb.records : [];
      const plan = MM.previewRecords(rows, keys(), overrides(), state.mode, state.row);
      const rowHtml = plan.rowRec ? await toHtml(plan.rowRec) : null;
      const extraHtml = plan.extraRec ? await toHtml(plan.extraRec) : null;
      if (token !== state.pvToken) return;
      err.innerHTML = '';
      rowBox.innerHTML = rowHtml == null ? '' : (state.mode === 'append' && plan.rowNumber ? `<div class="label">Excel record ${plan.rowNumber}</div>` : '') + `<div class="page">${rowHtml}</div>`;
      extraBox.innerHTML = extraHtml == null ? '' : `<div class="label extra">Additional website-entered record (added after all ${rows.length} Excel record(s))</div><div class="page">${extraHtml}</div>`;
    } catch (e) {
      if (token !== state.pvToken) return;
      rowBox.innerHTML = extraBox.innerHTML = '';
      err.innerHTML = `<div class="msg error" role="alert">Preview unavailable: ${esc(friendly(e))}</div>`;
      if (!e.user) console.error(e);
    } finally { if (token === state.pvToken) $('#pvSpin').hidden = true; }
  }

  /* ---------- generate ---------- */
  $('#generate').addEventListener('click', () => {
    clearMessages(); busy(true, 'Merging…'); $('#generate').disabled = true;
    setTimeout(() => {
      try {
        const records = MM.buildRecords(state.wb ? state.wb.records : [], keys(), overrides(), state.mode);
        const bytes = MM.mergeDocx(state.tplBuf, records);
        const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'merged_document.docx';
        document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
        message('ok', `Downloaded merged_document.docx with ${records.length} record(s).`);
      } catch (e) { message('error', friendly(e)); if (!e.user) console.error(e); }
      busy(false); updateGenerate();
    }, 20);
  });

  $('#reset').addEventListener('click', () => {
    Object.assign(state, { tplBuf: null, tplName: '', placeholders: [], wbBuf: null, wb: null, values: {}, mode: 'replace', row: 0 });
    document.querySelector('input[name=mode][value=replace]').checked = true;
    clearMessages(); tplZone.show(''); wbZone.show(''); afterInspect();
  });

  window.__mailMergeState = state; // handy for debugging in the browser console
})();
