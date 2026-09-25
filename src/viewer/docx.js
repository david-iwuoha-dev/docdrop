// Word documents: SuperDoc renders and edits the real DOCX (no HTML conversion).
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export class DocxView {
  constructor({ host, toolbarHost, onDirty, onError }) {
    this.host = host;
    this.toolbarHost = toolbarHost;
    this.onDirty = onDirty;
    this.onError = onError;
    this.sd = null;
    this.editing = false;
    this.results = [];
    this.resultIndex = -1;
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

  search(query) {
    this.results = [];
    this.resultIndex = -1;
    if (!query || !this.sd?.search) return { count: 0, index: -1 };
    try {
      this.results = this.sd.search(query) || [];
    } catch {
      this.results = [];
    }
    return this.step(1);
  }

  step(dir) {
    const n = this.results.length;
    if (!n) return { count: 0, index: -1 };
    this.resultIndex = (this.resultIndex + dir + n) % n;
    try { this.sd.goToSearchResult(this.results[this.resultIndex]); } catch { /* ignore */ }
    return { count: n, index: this.resultIndex };
  }

  destroy() {
    try { this.sd?.destroy(); } catch { /* ignore */ }
    this.sd = null;
    this.editing = false;
    this.results = [];
  }
}
