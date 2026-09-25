// Word documents: SuperDoc renders and edits the real DOCX (no HTML conversion).
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export class DocxView {
  constructor({ host, toolbarHost, onDirty, onError }) {
    this.host = host;
    this.toolbarHost = toolbarHost;
    this.onDirty = onDirty;
    this.onError = onError;
    this.sd = null;
    this.editor = null;
    this.editing = false;
    this.results = [];
    this.resultIndex = -1;
    this.query = '';
    this.useBrowserFind = false;
  }

  async load(buf, name) {
    const [{ SuperDoc }] = await Promise.all([import('superdoc'), import('superdoc/style.css')]);
    this.destroy();
    this.host.innerHTML = '';
    this.toolbarHost.innerHTML = '';
    const file = new File([buf], name, { type: DOCX_MIME });

    await new Promise((resolve, reject) => {
      let settled = false;
      const done = (fn, v) => { if (!settled) { settled = true; fn(v); } };
      const timer = setTimeout(() => done(reject, new Error('The document took too long to open.')), 90000);

      this.sd = new SuperDoc({
        selector: '#' + this.host.id,
        toolbar: '#' + this.toolbarHost.id,
        document: file,
        documentMode: 'viewing',
        telemetry: { enabled: false },
        onReady: () => { clearTimeout(timer); done(resolve); },
        onEditorCreate: ({ editor }) => { this.editor = editor; },
        onEditorUpdate: () => { if (this.editing) this.onDirty(); },
        onContentError: (e) => { clearTimeout(timer); done(reject, e?.error || new Error('This document could not be read.')); },
        onException: (e) => {
          const err = e?.error || e;
          if (!settled) { clearTimeout(timer); done(reject, err); }
          else this.onError(err);
        }
      });
    });
  }

  setEditing(on) {
    this.editing = on;
    this.sd?.setDocumentMode(on ? 'editing' : 'viewing');
    if (on) this.sd?.focus?.();
  }

  async exportDocx() {
    const out = await this.sd.export({ triggerDownload: false });
    const blob = Array.isArray(out) ? out[0] : out;
    if (!(blob instanceof Blob)) throw new Error('The document could not be exported.');
    return blob.type ? blob : new Blob([blob], { type: DOCX_MIME });
  }

  // Try a real PDF export first; fall back to the browser's print window.
  async print(name) {
    try {
      const out = await this.sd.export({ exportType: ['pdf'], triggerDownload: false });
      const blob = Array.isArray(out) ? out.find((b) => b?.type === 'application/pdf') : out;
      if (blob instanceof Blob && blob.type === 'application/pdf') {
        const url = URL.createObjectURL(blob);
        chrome.tabs.create({ url });
        setTimeout(() => URL.revokeObjectURL(url), 5 * 60000);
        return;
      }
    } catch { /* PDF export not available in this SuperDoc version */ }
    document.title = name;
    window.print();
  }

  // SuperDoc keeps its search tools in different places depending on the version.
  getEditor() {
    return this.editor || this.sd?.activeEditor || this.sd?.editor || null;
  }

  search(query) {
    this.results = [];
    this.resultIndex = -1;
    this.query = query;
    this.useBrowserFind = false;
    if (!query) return { count: 0, index: -1 };

    const ways = [
      () => this.sd?.search?.(query),
      () => this.getEditor()?.commands?.search?.(query)
    ];
    for (const way of ways) {
      try {
        const r = way();
        if (Array.isArray(r) && r.length) { this.results = r; break; }
      } catch { /* try the next way */ }
    }

    if (!this.results.length) {
      // Fall back to the browser's own find, starting from the top of the document.
      this.useBrowserFind = true;
      const sel = window.getSelection();
      const range = document.createRange();
      range.setStart(this.host, 0);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
      return this.browserFind(1);
    }
    return this.step(1);
  }

  step(dir) {
    if (this.useBrowserFind) return this.browserFind(dir);
    const n = this.results.length;
    if (!n) return { count: 0, index: -1 };
    this.resultIndex = (this.resultIndex + dir + n) % n;
    const hit = this.results[this.resultIndex];
    try {
      if (this.sd?.goToSearchResult) this.sd.goToSearchResult(hit);
      else this.getEditor()?.commands?.goToSearchResult?.(hit);
    } catch { /* ignore */ }
    return { count: n, index: this.resultIndex };
  }

  // The browser's own find, kept inside the document area.
  browserFind(dir) {
    const q = this.query;
    const text = (this.host.innerText || '').toLowerCase();
    const count = q ? text.split(q.toLowerCase()).length - 1 : 0;
    if (!count) return { count: 0, index: -1 };
    for (let i = 0; i < count + 2; i++) {
      if (!window.find(q, false, dir < 0, true, false, false, false)) break;
      const node = window.getSelection().anchorNode;
      if (node && this.host.contains(node)) {
        this.resultIndex = (this.resultIndex + dir + count) % count;
        return { count, index: this.resultIndex };
      }
    }
    return { count, index: Math.max(0, this.resultIndex) };
  }

  destroy() {
    try { this.sd?.destroy(); } catch { /* ignore */ }
    this.sd = null;
    this.editor = null;
    this.editing = false;
    this.results = [];
    this.resultIndex = -1;
    this.useBrowserFind = false;
  }
}