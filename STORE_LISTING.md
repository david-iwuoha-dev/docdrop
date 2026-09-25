# Chrome Web Store submission kit

Everything to copy into the Chrome Web Store Developer Dashboard for DocDrop.

## Before you submit

1. Register as a Chrome Web Store developer at https://chrome.google.com/webstore/devconsole (one-time registration fee; check the current amount on that page).
2. Build the upload file: `npm run package`. Upload `docdrop-vX.Y.Z.zip`.
3. Take screenshots (see below).
4. Make sure `PRIVACY.md` is pushed to GitHub, so you can link to it.

## Store listing tab

**Name:** DocDrop

**Summary (132 characters max):**
Open, edit and present Word, Excel and PowerPoint files right in Chrome. Private: your files never leave your computer.

**Category:** Productivity (Tools)

**Language:** English

**Description:**

DocDrop opens Office files in a Chrome tab, the same way Chrome opens PDFs. Turn it on, and Word, Excel and PowerPoint files open as soon as they finish downloading. You can also drag files onto the DocDrop tab or pick them from your computer.

Word documents
- True page layout, including headers, footers, tables and sections
- Edit and save as a real .docx file
- Search, print or save as PDF

Spreadsheets
- Opens .xlsx, .xls, .xlsm, .ods and .csv
- Shows cell colours, fonts, borders, merged cells and column widths
- Edit cells with a formula bar, insert and delete rows and columns, rename and add sheets
- Search across every sheet, zoom in and out
- Saves with your formatting kept

PowerPoint
- Opens .pptx, .ppt and .ppsx
- Full-screen slideshow with transitions and animations

Also
- Save back to the original file, or save a copy
- Recent files list
- Light and dark appearance

Private by design: no accounts, no uploads, no tracking, no ads. Everything happens inside the extension on your computer.

DocDrop is free and open source (AGPL-3.0). Source code: https://github.com/YOUR-USERNAME/docdrop

Created by Dave.

**Screenshots:** 1280 x 800 pixels (or 640 x 400), up to 5. Suggested:
1. A Word document open in DocDrop
2. The same document in Edit mode with the formatting toolbar
3. A colourful spreadsheet with the formula bar
4. A PowerPoint deck with slide thumbnails
5. The popup with the switch turned on

Tip: set the Chrome window to exactly 1280 x 800 before taking screenshots, or crop them afterwards.

**Small promo tile (optional):** 440 x 280 pixels. The DocDrop icon and name on the slate blue background (#1F3A4D) works well.

## Privacy practices tab

**Single purpose:**
DocDrop opens, edits and presents Word, Excel and PowerPoint files inside Chrome.

**Permission justifications:**

- **downloads:** Detects when a Word, Excel or PowerPoint file finishes downloading so DocDrop can open it in a tab, and shows a Save As window when the user saves an edited file.
- **storage:** Saves the user's settings (automatic opening on or off, light or dark appearance).
- **Host permission file:///\*:** Reads the file the user just downloaded from their own computer so it can be displayed. DocDrop does not access any website.

**Are you using remote code?** No. All code is bundled in the extension package.

**Data usage:** Tick nothing. DocDrop does not collect or transmit any user data.

Then tick the three certification boxes (no selling data, no unrelated use, no creditworthiness use).

**Privacy policy URL:**
https://github.com/YOUR-USERNAME/docdrop/blob/main/PRIVACY.md

## Distribution tab

- Visibility: Public (or Unlisted to share only by link at first)
- Regions: All regions

## After submitting

Review usually takes from a few days to a couple of weeks. If Google asks questions, the permission justifications above are the usual answers. To publish an update, raise the version number in `package.json` and `public/manifest.json`, run `npm run package`, and upload the new zip under **Package**.
