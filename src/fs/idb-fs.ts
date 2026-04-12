const DB_NAME = 'hwpx-fs';
const DB_VERSION = 1;
const STORE_NAME = 'files';

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      req.result.createObjectStore(STORE_NAME);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function tx(
  db: IDBDatabase,
  mode: IDBTransactionMode,
): IDBObjectStore {
  return db.transaction(STORE_NAME, mode).objectStore(STORE_NAME);
}

function req<T>(r: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}

function key(docId: string, path: string): string {
  return `${docId}/${path}`;
}

export async function writeFile(docId: string, path: string, data: Uint8Array | string): Promise<void> {
  const db = await openDb();
  await req(tx(db, 'readwrite').put(data, key(docId, path)));
  db.close();
}

export async function readFile(docId: string, path: string): Promise<Uint8Array | string> {
  const db = await openDb();
  const result = await req(tx(db, 'readonly').get(key(docId, path)));
  db.close();
  if (result === undefined) throw new Error(`File not found: ${path}`);
  return result;
}

export async function readFileAsString(docId: string, path: string): Promise<string> {
  const data = await readFile(docId, path);
  if (typeof data === 'string') return data;
  return new TextDecoder().decode(data);
}

export async function exists(docId: string, path: string): Promise<boolean> {
  const db = await openDb();
  const count = await req(tx(db, 'readonly').count(key(docId, path)));
  db.close();
  return count > 0;
}

export async function readDir(docId: string, prefix: string): Promise<string[]> {
  const db = await openDb();
  const store = tx(db, 'readonly');
  const lo = key(docId, prefix);
  const hi = lo + '\uffff';
  const keys = await req(store.getAllKeys(IDBKeyRange.bound(lo, hi)));
  db.close();
  const prefixLen = docId.length + 1;
  return (keys as string[]).map(k => k.substring(prefixLen));
}

export async function deleteFile(docId: string, path: string): Promise<void> {
  const db = await openDb();
  await req(tx(db, 'readwrite').delete(key(docId, path)));
  db.close();
}

export async function deleteAll(docId: string): Promise<void> {
  const db = await openDb();
  const store = tx(db, 'readwrite');
  const lo = `${docId}/`;
  const hi = lo + '\uffff';
  const keys = await req(store.getAllKeys(IDBKeyRange.bound(lo, hi)));
  for (const k of keys) {
    store.delete(k);
  }
  await new Promise<void>((resolve, reject) => {
    store.transaction.oncomplete = () => resolve();
    store.transaction.onerror = () => reject(store.transaction.error);
  });
  db.close();
}
