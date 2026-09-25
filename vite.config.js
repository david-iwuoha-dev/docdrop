import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';
import { existsSync, mkdirSync, copyFileSync } from 'node:fs';

const here = (p) => fileURLToPath(new URL(p, import.meta.url));

// Copies the licence and notice files of bundled libraries into dist/licenses,
// as their licences ask when they are redistributed.
const LICENSED = ['superdoc', 'pptx-vanilla-viewer', 'exceljs', 'xlsx', 'three'];
function copyLicenses() {
  return {
    name: 'docdrop-copy-licenses',
    closeBundle() {
      const out = here('./dist/licenses');
      mkdirSync(out, { recursive: true });
      for (const f of ['LICENSE', 'NOTICE.md']) {
        if (existsSync(here('./' + f))) copyFileSync(here('./' + f), `${out}/DocDrop-${f}`);
      }
      for (const pkg of LICENSED) {
        for (const f of ['LICENSE', 'LICENSE.md', 'LICENSE.txt', 'NOTICE', 'NOTICE.md']) {
          const src = here(`./node_modules/${pkg}/${f}`);
          if (existsSync(src)) copyFileSync(src, `${out}/${pkg}-${f}`);
        }
      }
    }
  };
}

// DocDrop is built as a Chrome extension (Manifest V3).
// Everything is bundled locally because extensions may not load remote code.
export default defineConfig({
  root: here('./src'),
  publicDir: here('./public'),
  base: './',
       resolve: {
       alias: {
         '@napi-rs/canvas': here('./src/shared/node-canvas-stub.js')
       }
     },
  build: {
    outDir: here('./dist'),
    emptyOutDir: true,
    target: 'chrome116',
    sourcemap: false,
    modulePreload: { polyfill: false },
    chunkSizeWarningLimit: 30000,
    rollupOptions: {
      input: {
        viewer: here('./src/viewer.html'),
        popup: here('./src/popup.html'),
        background: here('./src/background.js')
      },
      output: {
        entryFileNames: (chunk) =>
          chunk.name === 'background' ? 'background.js' : 'assets/[name]-[hash].js'
      }
    }
  },
  worker: { format: 'es' },
  plugins: [copyLicenses()]
});
