/**
 * 列データ結合ツール（VLOOKUP機能）
 * ① 基本リスト（マスタ）と ② 挿入先リストを読み込み、
 * キー列で照合して①の値を②に転記し CSV で出力する
 */
'use strict';

// ─── 状態 ──────────────────────────────────────────────────
let vlDataA = null;  // { headers:[], rows:[[]] } 基本リスト
let vlDataB = null;  // { headers:[], rows:[[]], fileName } 挿入先
let vlColMappings = [];

// ─── DOM ───────────────────────────────────────────────────
const vlDropA    = document.getElementById('vl-drop-a');
const vlDropB    = document.getElementById('vl-drop-b');
const vlInputA   = document.getElementById('vl-input-a');
const vlInputB   = document.getElementById('vl-input-b');
const vlFileA    = document.getElementById('vl-file-a');
const vlFileB    = document.getElementById('vl-file-b');
const vlMapping  = document.getElementById('vl-mapping');
const vlKeyA     = document.getElementById('vl-key-a');
const vlKeyB     = document.getElementById('vl-key-b');
const vlColList  = document.getElementById('vl-col-list');
const vlAddCol   = document.getElementById('vl-add-col');
const vlRunBtn   = document.getElementById('vl-run-btn');
const vlResult   = document.getElementById('vl-result');
const vlPreview  = document.getElementById('vl-preview-table');
const vlDlWrap   = document.getElementById('vl-dl-btn-wrap');
const vlResetBtn = document.getElementById('vl-reset-btn');
const optEncEl   = document.getElementById('opt-encoding');

// ─── ユーティリティ ────────────────────────────────────────
function escH(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function stripExt(name) { return name.replace(/\.[^/.]+$/, ''); }

function cellText(cell) {
  if (!cell) return '';
  if (cell.w !== undefined && cell.w !== null) return String(cell.w);
  if (cell.v !== undefined && cell.v !== null) return String(cell.v);
  return '';
}

function dlCsv(csvText, filename, enc) {
  let data;
  if (enc === 'utf8bom') {
    const bom  = new Uint8Array([0xEF, 0xBB, 0xBF]);
    const body = new TextEncoder().encode(csvText);
    data = new Uint8Array(bom.length + body.length);
    data.set(bom); data.set(body, bom.length);
  } else {
    data = new TextEncoder().encode(csvText);
  }
  const blob = new Blob([data], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), { href: url, download: filename });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ─── ドロップ＆クリック ────────────────────────────────────
function setupVlDrop(dropEl, inputEl, side) {
  dropEl.addEventListener('dragover',  e => { e.preventDefault(); dropEl.classList.add('drag-over'); });
  dropEl.addEventListener('dragleave', e => { if (!dropEl.contains(e.relatedTarget)) dropEl.classList.remove('drag-over'); });
  dropEl.addEventListener('drop', e => {
    e.preventDefault(); dropEl.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) loadVlFile(f, side);
  });
  dropEl.addEventListener('click', () => inputEl.click());
  dropEl.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); inputEl.click(); } });
  inputEl.addEventListener('change', e => {
    const f = e.target.files[0];
    if (f) loadVlFile(f, side);
    inputEl.value = '';
  });
}
setupVlDrop(vlDropA, vlInputA, 'A');
setupVlDrop(vlDropB, vlInputB, 'B');

// ─── Excel ファイル読み込み ────────────────────────────────
async function loadVlFile(file, side) {
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    alert('Excelファイル（.xlsx/.xls）を選択してください。');
    return;
  }
  try {
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array', cellText: true, cellDates: true, raw: false });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const data = parseSheetToTable(ws);
    data.fileName = file.name;
    data.baseName = stripExt(file.name);

    if (side === 'A') {
      vlDataA = data;
      renderVlFileInfo(vlFileA, file.name, data.headers.length, data.rows.length, 'a');
      vlDropA.classList.add('vl-loaded');
      renderVlHeaderHint('a', data.headers);
    } else {
      vlDataB = data;
      renderVlFileInfo(vlFileB, file.name, data.headers.length, data.rows.length, 'b');
      vlDropB.classList.add('vl-loaded');
      renderVlHeaderHint('b', data.headers);
    }

    if (vlDataA && vlDataB) buildMappingUI();
  } catch (e) {
    alert('ファイルの読み込みに失敗しました: ' + e.message);
  }
}

// シートをテーブル形式に変換
function parseSheetToTable(ws) {
  if (!ws['!ref']) return { headers: [], rows: [] };
  const range = XLSX.utils.decode_range(ws['!ref']);
  const headers = [];
  const rows = [];

  for (let C = range.s.c; C <= range.e.c; C++) {
    const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c: C })];
    headers.push(cellText(cell));
  }
  for (let R = range.s.r + 1; R <= range.e.r; R++) {
    const row = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      row.push(cellText(cell));
    }
    rows.push(row);
  }
  return { headers, rows };
}

// ファイル情報表示
function renderVlFileInfo(el, name, cols, rows, side) {
  el.classList.remove('hidden');
  const color = side === 'a' ? '#6366f1' : '#0891b2';
  el.innerHTML = `
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
      <polyline points="14 2 14 8 20 8"/>
    </svg>
    <strong>${escH(name)}</strong>
    <span class="file-meta">${cols} 列 / ${rows} 行</span>`;
}

// ヘッダーヒント表示
function renderVlHeaderHint(side, headers) {
  const infoEl = side === 'a' ? vlFileA : vlFileB;
  const existing = infoEl.querySelector('.header-hint');
  if (existing) existing.remove();

  const hint = document.createElement('div');
  hint.className = 'header-hint';
  hint.innerHTML = `<span class="hint-label">列名：</span>` +
    headers.map(h => `<span class="header-tag">${escH(h)}</span>`).join('');
  infoEl.appendChild(hint);

  const waitEl = document.getElementById(`vl-wait-${side}`);
  if (!waitEl) {
    const wait = document.createElement('p');
    wait.id = `vl-wait-${side}`;
    wait.className = 'wait-msg';
    const other = side === 'a' ? '②挿入先リスト' : '①基本リスト';
    const bothReady = vlDataA && vlDataB;
    if (!bothReady) {
      wait.textContent = `✅ 読み込み完了！次に ${other} を読み込んでください。`;
      infoEl.appendChild(wait);
    }
  }
}

// ─── マッピングUI 構築 ─────────────────────────────────────
function buildMappingUI() {
  vlMapping.classList.remove('hidden');
  vlResult.classList.add('hidden');
  document.querySelectorAll('.wait-msg').forEach(el => el.remove());

  vlKeyA.innerHTML = '<option value="">キー列を選択（① ' + escH(vlDataA.fileName) + '）</option>' +
    vlDataA.headers.map((h, i) => `<option value="${i}">${escH(h)}</option>`).join('');
  vlKeyB.innerHTML = '<option value="">キー列を選択（② ' + escH(vlDataB.fileName) + '）</option>' +
    vlDataB.headers.map((h, i) => `<option value="${i}">${escH(h)}</option>`).join('');

  vlColList.innerHTML = '';
  vlColMappings = [];
  addColMapping();
  checkRunnable();
}

// 転記列マッピング1行追加
function addColMapping() {
  const idx = vlColMappings.length;
  vlColMappings.push({ fromCol: '', toCol: '' });

  const row = document.createElement('div');
  row.className = 'col-row';
  row.dataset.idx = idx;
  row.innerHTML = `
    <div class="select-wrap">
      <span class="select-badge badge-a">①</span>
      <select class="vl-select from-col" data-idx="${idx}">
        <option value="">転記元の列（①）</option>
        ${vlDataA.headers.map((h, i) => `<option value="${i}">${escH(h)}</option>`).join('')}
      </select>
    </div>
    <div class="arrow-small">→</div>
    <div class="select-wrap">
      <span class="select-badge badge-b">②</span>
      <select class="vl-select to-col" data-idx="${idx}">
        <option value="">転記先の列（②）</option>
        ${vlDataB.headers.map((h, i) => `<option value="${i}">${escH(h)}</option>`).join('')}
      </select>
    </div>
    <button class="del-col-btn" data-idx="${idx}" title="この行を削除">✕</button>
  `;

  row.querySelector('.from-col').addEventListener('change', e => {
    vlColMappings[+e.target.dataset.idx].fromCol = e.target.value;
    checkRunnable();
  });
  row.querySelector('.to-col').addEventListener('change', e => {
    vlColMappings[+e.target.dataset.idx].toCol = e.target.value;
    checkRunnable();
  });
  row.querySelector('.del-col-btn').addEventListener('click', () => {
    row.remove();
    checkRunnable();
  });

  vlColList.appendChild(row);
}

vlAddCol.addEventListener('click', addColMapping);
vlKeyA.addEventListener('change', checkRunnable);
vlKeyB.addEventListener('change', checkRunnable);

function checkRunnable() {
  const keyOk = vlKeyA.value !== '' && vlKeyB.value !== '';
  const colRows = vlColList.querySelectorAll('.col-row');
  let colOk = false;
  colRows.forEach(row => {
    const from = row.querySelector('.from-col').value;
    const to   = row.querySelector('.to-col').value;
    if (from !== '' && to !== '') colOk = true;
  });
  vlRunBtn.disabled = !(keyOk && colOk);
}

// ─── 転記実行 ─────────────────────────────────────────────
vlRunBtn.addEventListener('click', runVlookup);

function runVlookup() {
  const keyAIdx = +vlKeyA.value;
  const keyBIdx = +vlKeyB.value;

  const colRows = vlColList.querySelectorAll('.col-row');
  const mappings = [];
  colRows.forEach(row => {
    const from = row.querySelector('.from-col').value;
    const to   = row.querySelector('.to-col').value;
    if (from !== '' && to !== '') mappings.push({ from: +from, to: +to });
  });

  if (!mappings.length) { alert('転記列を1つ以上設定してください。'); return; }

  const masterMap = new Map();
  vlDataA.rows.forEach(row => {
    const key = String(row[keyAIdx] ?? '').trim();
    if (key !== '') masterMap.set(key, row);
  });

  const newRows = vlDataB.rows.map(row => {
    const newRow = [...row];
    const key = String(row[keyBIdx] ?? '').trim();
    const masterRow = masterMap.get(key);
    if (masterRow) {
      mappings.forEach(({ from, to }) => {
        newRow[to] = masterRow[from] ?? '';
      });
    }
    return newRow;
  });

  // プレビュー表示
  const allRows = [vlDataB.headers, ...newRows];
  const previewRows = allRows.slice(0, 21);
  const trs = previewRows.map((row, ri) => {
    const tds = row.map(c => `<td>${escH(c)}</td>`).join('');
    return `<tr class="${ri === 0 ? 'preview-header-row' : ''}">${tds}</tr>`;
  });
  vlPreview.innerHTML = `<table class="preview-table"><tbody>${trs.join('')}</tbody></table>`;

  // CSV生成
  const csvLines = [vlDataB.headers, ...newRows].map(row =>
    row.map(v => {
      const str = String(v ?? '');
      return `"${str.replace(/"/g, '""')}"`;
    }).join(',')
  );
  const csvText = csvLines.join('\r\n');

  // ダウンロードボタン
  const enc = optEncEl ? optEncEl.value : 'utf8bom';
  const csvName = `${vlDataB.baseName}_転記済み.csv`;
  vlDlWrap.innerHTML = '';
  const dlBtn = document.createElement('button');
  dlBtn.className = 'dl-btn';
  dlBtn.innerHTML = `
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
      <polyline points="7 10 12 15 17 10"/>
      <line x1="12" y1="15" x2="12" y2="3"/>
    </svg>
    ${escH(csvName)} をダウンロード
  `;
  dlBtn.addEventListener('click', () => dlCsv(csvText, csvName, enc));
  vlDlWrap.appendChild(dlBtn);

  vlResult.classList.remove('hidden');
  vlResult.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── リセット ─────────────────────────────────────────────
vlResetBtn.addEventListener('click', () => {
  vlDataA = null; vlDataB = null;
  vlFileA.classList.add('hidden'); vlFileB.classList.add('hidden');
  vlDropA.classList.remove('vl-loaded'); vlDropB.classList.remove('vl-loaded');
  vlMapping.classList.add('hidden');
  vlResult.classList.add('hidden');
  vlColList.innerHTML = '';
  vlColMappings = [];
  document.querySelectorAll('.header-hint, .wait-msg').forEach(el => el.remove());
});
