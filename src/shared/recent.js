// Recently opened files, kept only on this computer (IndexedDB inside the extension).
const DB_NAME = 'docdrop';
const STORE = 'recent';
const MAX_ITEMS = 8;
const MAX_BYTES = 30 * 1024 * 1024;

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function run(mode, fn) {
  return openDb().then(
    (db) =>
      new Promise((resolve, reject) => {
        const tx = db.transaction(STORE, mode);
        const store = tx.objectStore(STORE);
        let result;
        Promise.resolve(fn(store)).then((r) => (result = r));
        tx.oncomplete = () => { db.close(); resolve(result); };
        tx.onerror = () => { db.close(); reject(tx.error); };
      })
  );
}

function all(store) {
  return new Promise((resolve, reject) => {
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

export async function addRecent(name, blob) {
  if (!blob || blob.size > MAX_BYTES) return;
  try {
    await run('readwrite', async (store) => {
      const items = await all(store);
      for (const it of items) {
        if (it.name === name && it.size === blob.size) store.delete(it.id);
      }
      const keep = items
        .filter((it) => !(it.name === name && it.size === blob.size))
        .sort((a, b) => b.openedAt - a.openedAt);
      keep.slice(MAX_ITEMS - 1).forEach((it) => store.delete(it.id));
      store.put({ id: crypto.randomUUID(), name, size: blob.size, openedAt: Date.now(), blob });
    });
  } catch (e) {
    console.warn('DocDrop: could not save recent file', e);
  }
}

export async function listRecent() {
  try {
    const items = await run('readonly', all);
    return items
      .sort((a, b) => b.openedAt - a.openedAt)
      .map(({ id, name, size, openedAt }) => ({ id, name, size, openedAt }));
  } catch {
    return [];
  }
}

export async function getRecent(id) {
  const items = await run('readonly', all);
  return items.find((it) => it.id === id) || null;
}

export async function touchRecent(id) {
  try {
    await run('readwrite', async (store) => {
      const items = await all(store);
      const it = items.find((x) => x.id === id);
      if (it) store.put({ ...it, openedAt: Date.now() });
    });
  } catch { /* not important */ }
}

export async function clearRecent() {
  await run('readwrite', (store) => store.clear());
}
