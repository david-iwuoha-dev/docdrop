// DocDrop: popup controls
const toggle = document.getElementById("toggle");
const status = document.getElementById("status");

function paint(on) {
  toggle.checked = on;
  status.textContent = on ? "DocDrop is on." : "DocDrop is off.";
  status.classList.toggle("on", on);
}

chrome.storage.local.get("enabled").then(({ enabled }) => paint(enabled === true));

toggle.addEventListener("change", async () => {
  paint(toggle.checked);
  await chrome.runtime.sendMessage({ type: "setEnabled", enabled: toggle.checked });
  if (toggle.checked) window.close();
});

document.getElementById("openTab").addEventListener("click", async () => {
  await chrome.runtime.sendMessage({ type: "openDropTab" });
  window.close();
});
