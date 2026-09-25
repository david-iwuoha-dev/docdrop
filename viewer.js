// DocDrop: renders .docx and spreadsheets in the tab
const $ = (id) => document.getElementById(id);
const SHEET_TYPES = ["xlsx", "xlsm", "xls", "csv", "ods"];

function extensionOf(name) {
  const m = /\.([a-z0-9]+)$/i.exec(name || "");
  return m ? m[1].toLowerCase() : "";
}
function baseName(path) {
  return (path || "").split(/[\\/]/).pop();
}

// Turn "C:\Users\me\Downloads\a b.docx" or "/Users/me/a.docx" into a file:/// URL
function toFileUrl(path) {
  const parts = path.replace(/\\/g, "/").split("/");
  const encoded = parts.map((p, i) => (i === 0 && /^[a-zA-Z]:$/.test(p)) ? p : encodeURIComponent(p));
  const joined = encoded.join("/");
  return "file://" + (joined.startsWith("/") ? "" : "/") + joined;
}

// Read a local file. Works when "Allow access to file URLs" is on.
function readLocal(path) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("GET", toFileUrl(path));
    xhr.responseType = "arraybuffer";
    xhr.onload = () => (xhr.response && xhr.response.byteLength)
      ? resolve(xhr.response) : reject(new Error("empty"));
    xhr.onerror = () => reject(new Error("blocked"));
    xhr.send();
  });
}

async function readRemote(url) {
  if (!/^https?:/i.test(url)) throw new Error("no url");
  const res = await fetch(url, { credentials: "include" });
  if (!res.ok) throw new Error("http " + res.status);
  return res.arrayBuffer();
}

function showMessage(lead, subHtml, isError) {
  $("doc").hidden = true; $("sheet").hidden = true; $("sheetTabs").hidden = true;
  const box = $("message");
  box.hidden = false;
  box.classList.toggle("error", !!isError);
  box.innerHTML = `<p class="lead"></p><div class="sub">${subHtml}</div>`;
  box.querySelector(".lead").textContent = lead;
}

function setHeader(name) {
  const ext = extensionOf(name);
  $("title").textContent = name;
  document.title = name;
  $("kind").textContent = ext.toUpperCase();
  $("kind").hidden = false;
}

async function renderDocx(buf) {
  if (!window.docx) throw new Error("The Word viewer library is missing from the lib folder.");
  const box = $("doc");
  box.innerHTML = "";
  box.hidden = false;
  await window.docx.renderAsync(buf, box, null, {
    inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: true
  });
}

function renderSheets(buf) {
  if (!window.XLSX) throw new Error("The spreadsheet library is missing from the lib folder.");
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const tabs = $("sheetTabs");
  tabs.innerHTML = "";
  tabs.setAttribute("role", "tablist");

  const show = (name) => {
    const html = XLSX.utils.sheet_to_html(wb.Sheets[name], { editable: false });
    const table = new DOMParser().parseFromString(html, "text/html").querySelector("table");
    const area = $("sheet");
    area.innerHTML = "";
    area.appendChild(table || document.createTextNode("This sheet is empty."));
    area.hidden = false;
    tabs.querySelectorAll("button").forEach(b =>
      b.setAttribute("aria-selected", String(b.dataset.name === name)));
    $("stage").scrollTo(0, 0);
  };

  wb.SheetNames.forEach(name => {
    const b = document.createElement("button");
    b.type = "button";
    b.role = "tab";
    b.textContent = name;
    b.dataset.name = name;
    b.onclick = () => show(name);
    tabs.appendChild(b);
  });
  tabs.hidden = wb.SheetNames.length < 2;
  show(wb.SheetNames[0]);
}

async function render(buf, name) {
  const ext = extensionOf(name);
  $("message").hidden = true;
  $("doc").hidden = true; $("sheet").hidden = true; $("sheetTabs").hidden = true;
  setHeader(name);
  try {
    if (ext === "docx") await renderDocx(buf);
    else if (SHEET_TYPES.includes(ext)) renderSheets(buf);
    else showMessage("This file type isn't supported.",
      "<p>DocDrop opens .docx, .xlsx, .xls, .csv and .ods files.</p>", true);
  } catch (err) {
    console.error(err);
    showMessage("This file couldn't be opened.",
      `<p>${err.message && err.message.includes("library") ? err.message : "It may be damaged, password protected, or an old .doc file."}</p>`, true);
  }
}

// Opened automatically after a download
async function loadFromParams() {
  const q = new URLSearchParams(location.search);
  const path = q.get("path");
  if (!path) return;
  const name = baseName(path);
  setHeader(name);
  showMessage("Opening " + name + "…", "", false);

  try { return render(await readLocal(path), name); } catch (_) {}
  try { return render(await readRemote(q.get("url") || ""), name); } catch (_) {}

  const allowed = await new Promise(r =>
    chrome.extension.isAllowedFileSchemeAccess ? chrome.extension.isAllowedFileSchemeAccess(r) : r(true));
  showMessage("Drag the file here to open it.",
    allowed
      ? "<p>The file couldn't be read automatically. Drag it in from Chrome's downloads list.</p>"
      : `<p>To open downloads automatically, turn on file access once:</p>
         <ol><li>Go to <code>chrome://extensions</code></li>
         <li>Click <b>Details</b> on DocDrop</li>
         <li>Turn on <b>Allow access to file URLs</b></li></ol>`,
    true);
}

// Drag and drop, plus the Open button
function openFile(file) {
  if (!file) return;
  file.arrayBuffer().then(buf => render(buf, file.name));
}
let depth = 0;
addEventListener("dragenter", e => { e.preventDefault(); depth++; $("dropCover").hidden = false; });
addEventListener("dragleave", () => { if (--depth <= 0) { depth = 0; $("dropCover").hidden = true; } });
addEventListener("dragover", e => e.preventDefault());
addEventListener("drop", e => {
  e.preventDefault(); depth = 0; $("dropCover").hidden = true;
  openFile(e.dataTransfer.files[0]);
});
$("openBtn").onclick = () => $("picker").click();
$("picker").onchange = e => openFile(e.target.files[0]);

loadFromParams();
