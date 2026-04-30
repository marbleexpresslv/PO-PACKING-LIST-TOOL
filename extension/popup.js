(function () {
  'use strict';

  const STORAGE_KEY = 'isf_parsed';

  function el(tag, attrs = {}, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') node.className = v;
      else if (k === 'style') node.style.cssText = v;
      else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
      else node.setAttribute(k, v);
    }
    for (const c of children) {
      if (c == null) continue;
      node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    }
    return node;
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  async function loadStored() {
    try {
      const data = await chrome.storage.session.get(STORAGE_KEY);
      return data[STORAGE_KEY] || null;
    } catch {
      return null;
    }
  }

  async function saveStored(parsed) {
    await chrome.storage.session.set({ [STORAGE_KEY]: parsed });
  }

  async function clearStored() {
    await chrome.storage.session.remove(STORAGE_KEY);
  }

  async function render() {
    const root = document.getElementById('root');
    root.innerHTML = '';
    const parsed = await loadStored();

    if (!parsed) {
      root.appendChild(el('p', {}, 'Carrega o Excel exportado pelo PO Packing List Tool.'));
      const input = el('input', { type: 'file', id: 'file', accept: '.xlsx,.xls' });
      input.addEventListener('change', handleFile);
      root.appendChild(input);
      root.appendChild(el('div', { class: 'hint' },
        'Depois de carregado, abre a modal "Packinglist Details" no Stone Profits, escolhe a unidade "m", clica em "Copy TSV" do bundle correto e cola na primeira célula da modal (Cmd+V).'
      ));
      return;
    }

    const s = parsed.summary;
    const meta = el('div', { id: 'meta' });
    meta.innerHTML = `
      <div class="meta-row"><b>Supplier:</b> ${escapeHtml(s.supplier || '—')}</div>
      <div class="meta-row"><b>PL:</b> ${escapeHtml(s.packingList || '—')} &nbsp; <b>PO:</b> ${escapeHtml(s.po || '—')}</div>
      <div class="meta-row"><b>Unit:</b> ${escapeHtml(s.unit || '—')}</div>
    `;
    root.appendChild(meta);

    const unit = (s.unit || '').toString().trim().toLowerCase();
    if (unit !== 'm') {
      root.appendChild(el('div', { id: 'unit-warn' },
        '⚠️ Unidade do Excel não é "m". Selecione "m" no toggle da modal antes de colar.'
      ));
    }

    const list = el('div', { id: 'bundles-list' });
    if (parsed.bundles.length === 0) {
      list.appendChild(el('div', { id: 'error' }, 'Nenhuma aba de BUNDLE encontrada no Excel.'));
    }
    for (const b of parsed.bundles) {
      const tsv = bundleToTsv(b);
      const row = el('div', { class: 'bundle' },
        el('div', { class: 'bundle-info' },
          el('span', { class: 'bundle-num' }, `BUNDLE ${b.number}`),
          el('span', { class: 'bundle-meta' }, `${b.rows.length} slabs`)
        )
      );
      const btn = el('button', { class: 'copy-btn' }, 'Copy TSV');
      btn.addEventListener('click', () => copyToClipboard(tsv, btn));
      row.appendChild(btn);
      list.appendChild(row);
    }
    root.appendChild(list);

    const clear = el('button', { id: 'clear' }, 'Carregar outro Excel');
    clear.addEventListener('click', async () => {
      await clearStored();
      render();
    });
    root.appendChild(clear);
  }

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
      const parsed = parseWorkbook(wb);
      await saveStored(parsed);
      render();
    } catch (err) {
      alert('Erro lendo Excel: ' + err.message);
      console.error(err);
    }
  }

  function parseWorkbook(wb) {
    const result = { summary: {}, bundles: [] };

    const summaryName = wb.SheetNames.find(n => n.toLowerCase().trim() === 'summary');
    if (summaryName) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[summaryName], { header: 1, defval: '' });
      for (const row of rows) {
        const k = (row[0] || '').toString().trim().toLowerCase();
        const v = row[1];
        if (k === 'supplier') result.summary.supplier = v;
        else if (k === 'packing list') result.summary.packingList = v;
        else if (k === 'po number') result.summary.po = v;
        else if (k === 'container') result.summary.container = v;
        else if (k === 'unit') result.summary.unit = v;
      }
    }

    for (const name of wb.SheetNames) {
      const m = name.match(/^bundle\s+(.+)$/i);
      if (!m) continue;
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
      const dataRows = rows.slice(1).filter(r => r.some(c => c !== '' && c != null));
      result.bundles.push({
        number: m[1].trim(),
        rows: dataRows.map(r => ({
          length: r[0],
          width: r[1],
          qty: r[2],
          block: r[3],
          bundle: r[4],
          ref: r[5],
        })),
      });
    }

    return result;
  }

  function bundleToTsv(b) {
    return b.rows.map(r => [
      fmt(r.length),
      fmt(r.width),
      fmt(r.qty),
      r.block ?? '',
      r.bundle ?? '',
      r.ref ?? '',
    ].join('\t')).join('\n');
  }

  function fmt(v) {
    if (v == null || v === '') return '';
    return v.toString();
  }

  async function copyToClipboard(text, btn) {
    try {
      await navigator.clipboard.writeText(text);
      const orig = btn.textContent;
      btn.textContent = '✓ Copiado!';
      btn.classList.add('copied');
      setTimeout(() => {
        btn.textContent = orig;
        btn.classList.remove('copied');
      }, 1500);
    } catch (err) {
      alert('Erro copiando: ' + err.message);
    }
  }

  render();
})();
