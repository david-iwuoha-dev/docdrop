// Spreadsheets: ExcelJS keeps formatting when saving; SheetJS reads older formats
// (.xls, .ods, .csv) and formats numbers and dates the way Excel shows them.
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

const MAX_CELLS = 250000;
const MAX_ROWS = 10000;
const MAX_COLS = 300;

const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

// Excel's legacy "indexed" colour palette.
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
];

const BORDER = {
  thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed',
  dashDot: '1px dashed', dashDotDot: '1px dashed', medium: '2px solid', mediumDashed: '2px dashed',
  mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed',
  thick: '3px solid', double: '3px double'
};

// ---------- small helpers ----------

export function colLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function decodeRef(ref) {
  const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref.trim());
  let c = 0;
  for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: Number(m[2]), c };
}

function esc(s) {
  return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function extOf(name) {
  return ((/\.([a-z0-9]+)$/i.exec(name) || [])[1] || '').toLowerCase();
}

function dateToSerial(d) {
  return (d.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
}

function formatNumber(n, fmt) {
  try {
    return XLSX.SSF.format(fmt || 'General', n);
  } catch {
    return String(n);
  }
}

function hexToRgb(h) {
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

function applyTint(hex, tint) {
  if (!tint) return hex;
  let [r, g, b] = hexToRgb(hex).map((v) => v / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0, l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  if (s === 0) r = g = b = l;
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    r = hue(p, q, h + 1 / 3); g = hue(p, q, h); b = hue(p, q, h - 1 / 3);
  }
  return [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function readTheme(wb) {
  try {
    const themes = wb._themes || {};
    const xml = themes.theme1 || Object.values(themes)[0];
    if (!xml) return DEFAULT_THEME;
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    return names.map((n, i) => {
      const el = doc.getElementsByTagNameNS('*', n)[0]?.firstElementChild;
      return el?.getAttribute('lastClr') || el?.getAttribute('val') || DEFAULT_THEME[i];
    });
  } catch {
    return DEFAULT_THEME;
  }
}

// ---------- cell values ----------

function richToText(v) {
  return Array.isArray(v?.richText) ? v.richText.map((t) => t.text).join('') : v;
}

export function displayText(cell) {
  if (!cell) return '';
  let v = cell.value;
  if (v == null) return '';
  if (typeof v === 'object' && !(v instanceof Date)) {
    if (v.formula || v.sharedFormula || 'result' in v) {
      v = v.result;
      if (v == null) return cell.formula ? '=' + cell.formula : '';
      if (typeof v === 'object' && v.error) return v.error;
    } else if (v.richText) return richToText(v);
    else if (v.hyperlink) return richToText(v.text) ?? v.hyperlink;
    else if (v.error) return v.error;
  }
  if (v instanceof Date) return formatNumber(dateToSerial(v), cell.numFmt || 'yyyy-mm-dd');
  if (typeof v === 'number') return formatNumber(v, cell.numFmt);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function rawText(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'object' && !(v instanceof Date)) {
    if (v.formula || v.sharedFormula) {
      const f = cell.formula || v.formula;
      return f ? '=' + f : displayText(cell);
    }
    return displayText(cell);
  }
  if (v instanceof Date) return formatNumber(dateToSerial(v), 'yyyy-mm-dd');
  if (typeof v === 'number') return String(v);
  return displayText(cell);
}

function isNumeric(cell) {
  if (!cell) return false;
  let v = cell.value;
  if (v && typeof v === 'object' && !(v instanceof Date) && 'result' in v) v = v.result;
  return typeof v === 'number' || v instanceof Date;
}

function parseInput(text, cell) {
  const t = text.trim();
  if (t === '') return { value: null };
  if (t.startsWith('=') && t.length > 1) return { value: { formula: t.slice(1) } };
  if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t)) return { value: Number(t) };
  if (/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return { value: Number(t.replace(/,/g, '')) };
  if (/^[-+]?(\d+\.?\d*|\.\d+)%$/.test(t)) return { value: Number(t.slice(0, -1)) / 100, numFmt: cell.numFmt || '0%' };
  if (/^(true|false)$/i.test(t)) return { value: t.toUpperCase() === 'TRUE' };
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) {
    const d = new Date(t + 'T00:00:00Z');
    if (!isNaN(d)) return { value: d, numFmt: cell.numFmt || 'yyyy-mm-dd' };
  }
  return { value: text };
}

function cloneValue(v) {
  try { return structuredClone(v); } catch { return v; }
}

// ---------- the view ----------

export class SheetView {
  constructor({ host, tabsEl, refEl, formulaEl, onDirty, onStatus }) {
    this.host = host;
    this.tabsEl = tabsEl;
    this.refEl = refEl;
    this.formulaEl = formulaEl;
    this.onDirty = onDirty;
    this.onStatus = onStatus;
    this.wb = null;
    this.active = 0;
    this.sel = { r: 1, c: 1 };
    this.editing = false;
    this.zoom = 1;
    this.undoStack = [];
    this.results = [];
    this.resultIndex = -1;
    this.bind();
  }

  // ----- loading -----

  async load(buf, name) {
    this.name = name;
    this.ext = extOf(name);
    this.convertedFrom = null;
    const wb = new ExcelJS.Workbook();

    const viaSheetJS = () => {
      const src = XLSX.read(buf, { type: 'array', cellDates: true, cellStyles: true });
      return XLSX.write(src, { bookType: 'xlsx', type: 'array' });
    };

    if (this.ext === 'xlsx' || this.ext === 'xlsm') {
      try {
        await wb.xlsx.load(buf);
      } catch (err) {
        console.warn('DocDrop: ExcelJS could not read this file, using SheetJS instead', err);
        await wb.xlsx.load(viaSheetJS());
        this.onStatus('Some formatting could not be read from this file.');
      }
      if (this.ext === 'xlsm') this.convertedFrom = 'xlsm';
    } else {
      await wb.xlsx.load(viaSheetJS());
      if (this.ext !== 'csv') this.convertedFrom = this.ext;
    }

    if (!wb.worksheets.length) throw new Error('This workbook has no sheets.');
    this.wb = wb;
    this.theme = readTheme(wb);
    this.undoStack = [];
    const firstVisible = wb.worksheets.findIndex((ws) => !ws.state || ws.state === 'visible');
    this.active = Math.max(0, firstVisible);
    this.sel = { r: 1, c: 1 };
    this.renderTabs();
    this.render();
  }

  get ws() {
    return this.wb.worksheets[this.active];
  }

  // ----- styles -----

  color(c) {
    if (!c) return null;
    if (c.argb) return '#' + c.argb.slice(-6);
    if (c.theme != null) {
      const base = this.theme[c.theme];
      return base ? '#' + applyTint(base, c.tint) : null;
    }
    if (c.indexed != null) return INDEXED[c.indexed] ? '#' + INDEXED[c.indexed] : null;
    return null;
  }

  cellStyle(cell) {
    const css = [];
    const f = cell.font;
    if (f) {
      if (f.bold) css.push('font-weight:700');
      if (f.italic) css.push('font-style:italic');
      const deco = [f.underline && 'underline', f.strike && 'line-through'].filter(Boolean).join(' ');
      if (deco) css.push('text-decoration:' + deco);
      if (f.size) css.push(`font-size:${f.size}pt`);
      if (f.name) css.push(`font-family:"${f.name.replace(/"/g, '')}",Calibri,Arial,sans-serif`);
      const fc = this.color(f.color);
      if (fc) css.push('color:' + fc);
    }
    const fill = cell.fill;
    if (fill && fill.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
      const bg = this.color(fill.fgColor) || this.color(fill.bgColor);
      if (bg) css.push('background:' + bg);
    } else if (fill && fill.type === 'gradient' && fill.stops?.length) {
      const bg = this.color(fill.stops[0].color);
      if (bg) css.push('background:' + bg);
    }
    const a = cell.alignment;
    if (a) {
      if (a.horizontal && ['left', 'center', 'right', 'justify'].includes(a.horizontal)) css.push('text-align:' + a.horizontal);
      if (a.horizontal === 'centerContinuous') css.push('text-align:center');
      if (a.vertical) css.push('vertical-align:' + ({ middle: 'middle', top: 'top', bottom: 'bottom' }[a.vertical] || 'bottom'));
      if (a.wrapText) css.push('white-space:pre-wrap');
      if (a.indent) css.push(`padding-left:${4 + a.indent * 9}px`);
    }
    const b = cell.border;
    if (b) {
      for (const side of ['top', 'right', 'bottom', 'left']) {
        const s = b[side];
        if (s && s.style && BORDER[s.style]) css.push(`border-${side}:${BORDER[s.style]} ${this.color(s.color) || '#000'}`);
      }
    }
    return css.join(';');
  }

  // ----- rendering -----

  merges(ws) {
    const list = [];
    try {
      if (ws._merges) Object.values(ws._merges).forEach((m) => list.push(m.model || m));
      else (ws.model.merges || []).forEach((ref) => {
        const [a, b] = ref.split(':').map(decodeRef);
        list.push({ top: a.r, left: a.c, bottom: b.r, right: b.c });
      });
    } catch { /* no merges */ }
    const masters = new Map();
    const covered = new Set();
    for (const m of list) {
      masters.set(m.top + ',' + m.left, { rs: m.bottom - m.top + 1, cs: m.right - m.left + 1 });
      for (let r = m.top; r <= m.bottom; r++)
        for (let c = m.left; c <= m.right; c++)
          if (r !== m.top || c !== m.left) covered.add(r + ',' + c);
    }
    this.mergeMasters = masters;
    this.mergeCovered = covered;
    this.mergeList = list;
  }

  render() {
    const ws = this.ws;
    this.merges(ws);
    let rows = Math.max(ws.rowCount || 0, ws.actualRowCount || 0, 1);
    let cols = Math.max(ws.columnCount || 0, ws.actualColumnCount || 0, 1);
    for (const m of this.mergeList) { rows = Math.max(rows, m.bottom); cols = Math.max(cols, m.right); }
    // A little empty space to type into, like a real spreadsheet.
    if (this.editing) { rows += 20; cols += 5; }
    cols = Math.min(cols, MAX_COLS);
    rows = Math.min(rows, MAX_ROWS, Math.floor(MAX_CELLS / cols));
    this.rows = rows;
    this.cols = cols;
    const truncated = rows < (ws.rowCount || 0) || cols < (ws.columnCount || 0);

    const defaultW = ws.properties?.defaultColWidth || 8.43;
    const colgroup = ['<col class="rh">'];
    const headCells = ['<th class="corner"></th>'];
    for (let c = 1; c <= cols; c++) {
      const col = ws.getColumn(c);
      const w = col.hidden ? 0 : Math.round((col.width || defaultW) * 7 + 5);
      colgroup.push(`<col style="width:${w}px">`);
      headCells.push(`<th class="ch${col.hidden ? ' hid' : ''}" data-c="${c}">${colLetter(c)}</th>`);
    }

    const body = [];
    for (let r = 1; r <= rows; r++) {
      const row = ws.findRow(r); // findRow/findCell never create empty rows in the file
      const h = row?.height ? ` style="height:${Math.round(row.height * 4 / 3)}px"` : '';
      const cells = [`<th class="rn" data-r="${r}">${r}</th>`];
      for (let c = 1; c <= cols; c++) {
        const key = r + ',' + c;
        if (this.mergeCovered.has(key)) continue;
        const cell = row?.findCell(c);
        const m = this.mergeMasters.get(key);
        const span = m ? `${m.rs > 1 ? ` rowspan="${Math.min(m.rs, rows - r + 1)}"` : ''}${m.cs > 1 ? ` colspan="${Math.min(m.cs, cols - c + 1)}"` : ''}` : '';
        if (!cell) { cells.push(`<td data-r="${r}" data-c="${c}"${span}></td>`); continue; }
        const style = cell.style && Object.keys(cell.style).length ? this.cellStyle(cell) : '';
        const cls = isNumeric(cell) ? ' class="num"' : '';
        cells.push(`<td data-r="${r}" data-c="${c}"${span}${cls}${style ? ` style="${esc(style)}"` : ''}>${esc(displayText(cell))}</td>`);
      }
      body.push(`<tr${h}${row?.hidden ? ' class="hid"' : ''}>${cells.join('')}</tr>`);
    }

    this.host.innerHTML =
      `<div class="grid-wrap" tabindex="0"><table class="grid"><colgroup>${colgroup.join('')}</colgroup>` +
      `<thead><tr>${headCells.join('')}</tr></thead><tbody>${body.join('')}</tbody></table></div>`;
    this.wrap = this.host.firstElementChild;
    this.table = this.wrap.firstElementChild;
    this.table.style.zoom = this.zoom;
    if (truncated) this.onStatus(`Large sheet: showing the first ${rows.toLocaleString()} rows and ${cols} columns.`);
    this.select(Math.min(this.sel.r, rows), Math.min(this.sel.c, cols), false);
    this.highlightTabs();
  }

  renderTabs() {
    this.tabsEl.innerHTML = '';
    this.wb.worksheets.forEach((ws, i) => {
      if (ws.state && ws.state !== 'visible') return;
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'sheet-tab';
      b.textContent = ws.name;
      b.dataset.i = i;
      b.setAttribute('role', 'tab');
      b.title = 'Double-click to rename while editing';
      b.onclick = () => this.show(i);
      b.ondblclick = () => this.rename(i);
      this.tabsEl.appendChild(b);
    });
    if (this.editing) {
      const add = document.createElement('button');
      add.type = 'button';
      add.className = 'sheet-tab add';
      add.textContent = '+';
      add.title = 'Add a sheet';
      add.setAttribute('aria-label', 'Add a sheet');
      add.onclick = () => this.addSheet();
      this.tabsEl.appendChild(add);
    }
    this.highlightTabs();
  }

  highlightTabs() {
    this.tabsEl.querySelectorAll('.sheet-tab[data-i]').forEach((b) =>
      b.setAttribute('aria-selected', String(Number(b.dataset.i) === this.active)));
  }

  show(i) {
    if (i === this.active) return;
    this.active = i;
    this.sel = { r: 1, c: 1 };
    this.render();
  }

  tdAt(r, c) {
    return this.table?.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
  }

  masterOf(r, c) {
    if (!this.mergeCovered?.has(r + ',' + c)) return { r, c };
    const m = this.mergeList.find((m) => r >= m.top && r <= m.bottom && c >= m.left && c <= m.right);
    return m ? { r: m.top, c: m.left } : { r, c };
  }

  select(r, c, scroll = true) {
    ({ r, c } = this.masterOf(r, c));
    this.sel = { r, c };
    this.table.querySelectorAll('.sel').forEach((el) => el.classList.remove('sel'));
    this.table.querySelectorAll('.hl').forEach((el) => el.classList.remove('hl'));
    const td = this.tdAt(r, c);
    if (td) {
      td.classList.add('sel');
      if (scroll) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    this.table.querySelector(`th.ch[data-c="${c}"]`)?.classList.add('hl');
    this.table.querySelector(`th.rn[data-r="${r}"]`)?.classList.add('hl');
    this.refEl.textContent = colLetter(c) + r;
    this.formulaEl.value = rawText(this.peek(r, c));
  }

  // Read a cell without creating it.
  peek(r, c) {
    return this.ws.findRow(r)?.findCell(c) || null;
  }

  move(dr, dc) {
    let { r, c } = this.sel;
    const m = this.mergeMasters.get(r + ',' + c);
    if (m && dr > 0) r += m.rs - 1;
    if (m && dc > 0) c += m.cs - 1;
    r = Math.min(Math.max(1, r + dr), this.rows);
    c = Math.min(Math.max(1, c + dc), this.cols);
    this.select(r, c);
  }

  refreshCell(r, c) {
    const td = this.tdAt(r, c);
    if (!td) return;
    const cell = this.ws.getCell(r, c);
    td.textContent = displayText(cell);
    td.classList.toggle('num', isNumeric(cell));
  }

  // ----- editing -----

  setEditing(on) {
    this.editing = on;
    this.formulaEl.readOnly = !on;
    this.renderTabs();
    this.render();
  }

  commit(text, { r, c } = this.sel) {
    if (!this.editing) return;
    if (text === rawText(this.peek(r, c))) return;
    const cell = this.ws.getCell(r, c);
    this.undoStack.push({ sheet: this.active, r, c, value: cloneValue(cell.value), numFmt: cell.numFmt });
    if (this.undoStack.length > 200) this.undoStack.shift();
    const parsed = parseInput(text, cell);
    cell.value = parsed.value;
    if (parsed.numFmt) cell.numFmt = parsed.numFmt;
    this.refreshCell(r, c);
    this.onDirty();
  }

  undo() {
    const last = this.undoStack.pop();
    if (!last) return this.onStatus('Nothing to undo.');
    if (last.sheet !== this.active) this.show(last.sheet);
    const cell = this.ws.getCell(last.r, last.c);
    cell.value = last.value;
    cell.numFmt = last.numFmt;
    this.refreshCell(last.r, last.c);
    this.select(last.r, last.c);
    this.onDirty();
  }

  structural(fn, message) {
    if (!this.editing) return;
    fn(this.ws);
    this.undoStack = [];
    this.render();
    this.onDirty();
    this.onStatus(message);
  }

  insertRow() { const r = this.sel.r; this.structural((ws) => ws.spliceRows(r + 1, 0, []), `Row inserted below row ${r}.`); }
  deleteRow() { const r = this.sel.r; this.structural((ws) => ws.spliceRows(r, 1), `Row ${r} deleted.`); }
  insertCol() { const c = this.sel.c; this.structural((ws) => ws.spliceColumns(c + 1, 0, []), `Column inserted after ${colLetter(c)}.`); }
  deleteCol() { const c = this.sel.c; this.structural((ws) => ws.spliceColumns(c, 1), `Column ${colLetter(c)} deleted.`); }

  rename(i) {
    if (!this.editing) return this.onStatus('Turn on Edit to rename sheets.');
    const ws = this.wb.worksheets[i];
    const name = prompt('Rename sheet', ws.name);
    if (name == null) return;
    const clean = name.trim();
    if (!clean || clean.length > 31 || /[\\/?*[\]:]/.test(clean)) {
      return this.onStatus('Sheet names need 1 to 31 characters and cannot use \\ / ? * [ ] :');
    }
    if (this.wb.worksheets.some((w, j) => j !== i && w.name.toLowerCase() === clean.toLowerCase())) {
      return this.onStatus('Another sheet already has that name.');
    }
    ws.name = clean;
    this.renderTabs();
    this.onDirty();
  }

  addSheet() {
    let n = this.wb.worksheets.length + 1;
    while (this.wb.getWorksheet('Sheet' + n)) n++;
    this.wb.addWorksheet('Sheet' + n);
    this.active = this.wb.worksheets.length - 1;
    this.sel = { r: 1, c: 1 };
    this.renderTabs();
    this.render();
    this.onDirty();
  }

  // ----- search across every sheet -----

  search(query) {
    this.results = [];
    this.resultIndex = -1;
    const q = (query || '').trim().toLowerCase();
    if (!q) return { count: 0, index: -1 };
    this.wb.worksheets.forEach((ws, si) => {
      if (ws.state && ws.state !== 'visible') return;
      ws.eachRow({ includeEmpty: false }, (row, r) => {
        row.eachCell({ includeEmpty: false }, (cell, c) => {
          if (this.results.length < 2000 && displayText(cell).toLowerCase().includes(q)) this.results.push({ si, r, c });
        });
      });
    });
    return this.step(1);
  }

  step(dir) {
    const n = this.results.length;
    if (!n) return { count: 0, index: -1 };
    this.resultIndex = (this.resultIndex + dir + n) % n;
    const hit = this.results[this.resultIndex];
    if (hit.si !== this.active) { this.active = hit.si; this.render(); }
    if (hit.r > this.rows || hit.c > this.cols) {
      this.onStatus(`Match at ${colLetter(hit.c)}${hit.r} is beyond the part of the sheet shown.`);
    } else {
      this.select(hit.r, hit.c);
    }
    return { count: n, index: this.resultIndex, where: `${this.wb.worksheets[hit.si].name}!${colLetter(hit.c)}${hit.r}` };
  }

  // ----- zoom -----

  setZoom(z) {
    this.zoom = Math.min(3, Math.max(0.3, Math.round(z * 10) / 10));
    if (this.table) this.table.style.zoom = this.zoom;
    return this.zoom;
  }

  // ----- saving -----

  outputName() {
    if (this.ext === 'csv' || this.ext === 'xlsx') return this.name;
    return this.name.replace(/\.[^.]+$/, '') + '.xlsx';
  }

  async exportBlob() {
    if (this.ext === 'csv') {
      const ws = this.wb.worksheets[0];
      const lines = [];
      const cols = ws.columnCount;
      ws.eachRow({ includeEmpty: true }, (row, r) => {
        while (lines.length < r - 1) lines.push('');
        const out = [];
        for (let c = 1; c <= cols; c++) {
          const t = rawText(row.findCell(c));
          out.push(/[",\n\r]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t);
        }
        lines.push(out.join(','));
      });
      return new Blob([lines.join('\r\n')], { type: 'text/csv' });
    }
    this.wb.calcProperties = { ...(this.wb.calcProperties || {}), fullCalcOnLoad: true };
    const buf = await this.wb.xlsx.writeBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // ----- keyboard and mouse -----

  bind() {
    this.host.addEventListener('mousedown', (e) => {
      const td = e.target.closest('td[data-r]');
      if (!td) return;
      // Keep what was being typed in the formula bar before moving.
      if (this.editing && document.activeElement === this.formulaEl) this.commit(this.formulaEl.value);
      this.select(Number(td.dataset.r), Number(td.dataset.c), false);
    });
    this.host.addEventListener('dblclick', (e) => {
      if (e.target.closest('td[data-r]') && this.editing) this.startEdit();
    });
    this.host.addEventListener('keydown', (e) => {
      if (!this.table || e.target !== this.wrap) return;
      const keys = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (keys[e.key]) { e.preventDefault(); return this.move(...keys[e.key]); }
      if (e.key === 'Tab') { e.preventDefault(); return this.move(0, e.shiftKey ? -1 : 1); }
      if (e.key === 'Enter') { e.preventDefault(); return this.editing ? this.startEdit() : this.move(1, 0); }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); return this.editing && this.undo(); }
      if (!this.editing) return;
      if (e.key === 'F2') { e.preventDefault(); return this.startEdit(); }
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); this.commit(''); return this.select(this.sel.r, this.sel.c); }
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        this.startEdit(e.key);
      }
    });

    this.formulaEl.addEventListener('keydown', (e) => {
      if (!this.editing || !this.table) return;
      if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        this.commit(this.formulaEl.value);
        this.wrap.focus();
        if (e.key === 'Enter') this.move(e.shiftKey ? -1 : 1, 0);
        else this.move(0, e.shiftKey ? -1 : 1);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.select(this.sel.r, this.sel.c);
        this.wrap.focus();
      }
    });
    this.formulaEl.addEventListener('blur', () => {
      if (this.editing && this.table && this.formulaEl.value !== rawText(this.peek(this.sel.r, this.sel.c))) {
        this.commit(this.formulaEl.value);
      }
    });
  }

  startEdit(firstChar) {
    this.formulaEl.focus();
    if (firstChar != null) this.formulaEl.value = firstChar;
    const n = this.formulaEl.value.length;
    this.formulaEl.setSelectionRange(n, n);
  }

  destroy() {
    this.host.innerHTML = '';
    this.tabsEl.innerHTML = '';
    this.wb = null;
    this.table = null;
    this.results = [];
    this.undoStack = [];
  }
}
