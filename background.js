// DocDrop: background worker
const SUPPORTED = ["docx", "xlsx", "xlsm", "xls", "csv", "ods"];

async function isOn() {
  const { enabled } = await chrome.storage.local.get("enabled");
  return enabled === true;
}

async function showBadge(on) {
  await chrome.action.setBadgeText({ text: on ? "ON" : "OFF" });
  await chrome.action.setBadgeBackgroundColor({ color: on ? "#2E8B7A" : "#8A949B" });
}

chrome.runtime.onInstalled.addListener(async () => showBadge(await isOn()));
chrome.runtime.onStartup.addListener(async () => showBadge(await isOn()));

// Click the icon to toggle on/off
chrome.action.onClicked.addListener(async () => {
  const next = !(await isOn());
  await chrome.storage.local.set({ enabled: next });
  await showBadge(next);
});

function extensionOf(name) {
  const m = /\.([a-z0-9]+)$/i.exec(name || "");
  return m ? m[1].toLowerCase() : "";
}

// When a download finishes, open it in the viewer
chrome.downloads.onChanged.addListener(async (delta) => {
  if (!delta.state || delta.state.current !== "complete") return;
  if (!(await isOn())) return;

  const [item] = await chrome.downloads.search({ id: delta.id });
  if (!item || !SUPPORTED.includes(extensionOf(item.filename))) return;

  const params = new URLSearchParams({
    path: item.filename,
    url: item.finalUrl || item.url || ""
  });
  chrome.tabs.create({ url: chrome.runtime.getURL("viewer.html") + "?" + params.toString() });
});
