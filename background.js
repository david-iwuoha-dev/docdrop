// DocDrop: background worker
const SUPPORTED = ["docx", "xlsx", "xlsm", "xls", "csv", "ods"];
const VIEWER = chrome.runtime.getURL("viewer.html");

async function isOn() {
  const { enabled } = await chrome.storage.local.get("enabled");
  return enabled === true;
}

async function showBadge(on) {
  await chrome.action.setBadgeText({ text: on ? "ON" : "OFF" });
  await chrome.action.setBadgeBackgroundColor({ color: on ? "#2E8B7A" : "#8A949B" });
}

// Open the DocDrop tab, or switch to it if it's already open
async function openDropTab() {
  let blank = null;
  if (chrome.runtime.getContexts) {
    const pages = await chrome.runtime.getContexts({ contextTypes: ["TAB"] });
    blank = pages.find(p => p.documentUrl === VIEWER);
  }
  if (blank) {
    await chrome.tabs.update(blank.tabId, { active: true });
    await chrome.windows.update(blank.windowId, { focused: true });
  } else {
    await chrome.tabs.create({ url: VIEWER });
  }
}

chrome.runtime.onInstalled.addListener(async () => showBadge(await isOn()));
chrome.runtime.onStartup.addListener(async () => showBadge(await isOn()));

// Messages from the popup
chrome.runtime.onMessage.addListener((msg, _sender, reply) => {
  (async () => {
    if (msg.type === "setEnabled") {
      await chrome.storage.local.set({ enabled: msg.enabled });
      await showBadge(msg.enabled);
      if (msg.enabled) await openDropTab();
    } else if (msg.type === "openDropTab") {
      await openDropTab();
    }
    reply({ ok: true });
  })();
  return true;
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
  chrome.tabs.create({ url: VIEWER + "?" + params.toString() });
});
