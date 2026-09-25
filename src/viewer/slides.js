// PowerPoint: view and present only (pptx-vanilla-viewer, Apache-2.0).
const HIDDEN = [
  'share', 'broadcast', 'record', 'export', 'undo', 'redo',
  'file', 'home', 'insert', 'draw', 'design', 'transitions', 'animations', 'review'
];

export class SlidesView {
  constructor({ host, onError }) {
    this.host = host;
    this.onError = onError;
    this.viewer = null;
  }

  async load(buf, name) {
    const { createPptxViewer } = await import('pptx-vanilla-viewer');
    this.destroy();
    this.host.innerHTML = '';
    await new Promise((resolve, reject) => {
      let settled = false;
      const timer = setTimeout(() => { if (!settled) { settled = true; reject(new Error('The presentation took too long to open.')); } }, 90000);
      this.viewer = createPptxViewer(this.host, {
        source: buf,
        fileName: name,
        editable: false,
        showToolbar: true,
        showThumbnails: true,
        hiddenActions: HIDDEN,
        theme: { colors: { primary: '#2E8B7A' } },
        onLoad: () => { if (!settled) { settled = true; clearTimeout(timer); resolve(); } },
        onError: (message) => {
          const err = new Error(message || 'This presentation could not be opened.');
          if (!settled) { settled = true; clearTimeout(timer); reject(err); }
          else this.onError(err);
        }
      });
    });
  }

  present() {
    return this.viewer?.enterPresentation();
  }

  print() {
    const ok = this.viewer?.print();
    if (ok === false) window.print();
  }

  destroy() {
    try { this.viewer?.destroy(); } catch { /* ignore */ }
    this.viewer = null;
  }
}
