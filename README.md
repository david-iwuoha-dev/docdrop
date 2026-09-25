# DocDrop

A Chrome extension that opens Word documents and spreadsheets right in a browser tab, the same way Chrome opens PDFs.

When DocDrop is **ON**, any supported file you download opens automatically in a new tab as soon as the download finishes. You can also drag a file onto the DocDrop tab or use its **Open a file** button.

## Supported files

| Type | Extensions |
|---|---|
| Word | `.docx` |
| Spreadsheets | `.xlsx`, `.xlsm`, `.xls`, `.ods`, `.csv` |

Spreadsheets with several sheets get a tab bar at the bottom so you can switch between them.

## Setup

1. **Add the libraries.** Download these two files into the `lib/` folder (keep the names exactly as shown):
   - `docx-preview.min.js` from https://cdn.jsdelivr.net/npm/docx-preview@0.3.5/dist/docx-preview.min.js
   - `xlsx.full.min.js` from https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js

   In a terminal (e.g. in Codespaces) you can run:
   ```bash
   curl -L -o lib/docx-preview.min.js https://cdn.jsdelivr.net/npm/docx-preview@0.3.5/dist/docx-preview.min.js
   curl -L -o lib/xlsx.full.min.js https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
   ```
   `jszip.min.js` is already included.

2. **Load it into Chrome.** Go to `chrome://extensions`, turn on **Developer mode**, click **Load unpacked**, and select the `docdrop` folder.

3. **Allow file access.** Click **Details** on DocDrop and turn on **Allow access to file URLs**. This lets downloads open instantly.

4. **Turn it on.** Pin DocDrop to your toolbar and click its icon. The badge shows **ON** (green) or **OFF** (grey).

## Project structure

```
docdrop/
├── manifest.json     Extension settings and permissions (Manifest V3)
├── background.js     On/off toggle and download watcher
├── viewer.html       The viewer tab
├── viewer.css        Viewer styles
├── viewer.js         Reads and renders files
├── icons/            Toolbar and store icons
└── lib/              Rendering libraries (JSZip, docx-preview, SheetJS)
```

## How it works

- `background.js` listens for finished downloads. If DocDrop is ON and the file type is supported, it opens `viewer.html` with the file's path.
- `viewer.js` reads the file from disk (or re-fetches it from its original link if that fails), then renders it with **docx-preview** for Word files or **SheetJS** for spreadsheets.

## Limitations

- View only, no editing.
- Old `.doc` files (Word 97-2003) aren't supported.
- Complex Word layouts (text boxes, embedded charts) may look slightly different from Microsoft Word.
- Files downloaded while DocDrop is OFF won't open automatically; drag them in instead.

## Credits

- [JSZip](https://stuk.github.io/jszip/) (MIT)
- [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) (Apache 2.0)
- [SheetJS Community Edition](https://sheetjs.com/) (Apache 2.0)
