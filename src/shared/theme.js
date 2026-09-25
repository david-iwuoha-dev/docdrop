// Light / dark / automatic appearance, shared by the viewer and the popup.
const ORDER = ['light', 'dark', 'auto'];

export async function getTheme() {
  const { theme } = await chrome.storage.local.get('theme');
  return ORDER.includes(theme) ? theme : 'light';
}

export function applyTheme(theme) {
  const root = document.documentElement;
  if (theme === 'auto') root.removeAttribute('data-theme');
  else root.setAttribute('data-theme', theme);
}

export async function cycleTheme() {
  const next = ORDER[(ORDER.indexOf(await getTheme()) + 1) % ORDER.length];
  await chrome.storage.local.set({ theme: next });
  applyTheme(next);
  return next;
}

export function themeLabel(theme) {
  return { auto: 'Auto', dark: 'Dark', light: 'Light' }[theme];
}

export async function initTheme() {
  const t = await getTheme();
  applyTheme(t);
  chrome.storage.onChanged.addListener((changes) => {
    if (changes.theme) applyTheme(changes.theme.newValue || 'auto');
  });
  return t;
}
