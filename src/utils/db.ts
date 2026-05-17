export const DB_NAME = "DataControllerDB";
export const STORE_NAME = "keyval";

let dbInstance: IDBDatabase | null = null;
let initPromise: Promise<IDBDatabase> | null = null;

export function initDB(): Promise<IDBDatabase> {
  if (dbInstance) return Promise.resolve(dbInstance);
  if (initPromise) return initPromise;

  initPromise = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(STORE_NAME);
    };
    request.onsuccess = () => {
      dbInstance = request.result;
      resolve(dbInstance);
    };
    request.onerror = () => reject(request.error);
  });
  
  return initPromise;
}

const pendingGets = new Map<string, Promise<string | null>>();

export async function idbGet(key: string): Promise<string | null> {
  // Coalesce identical concurrent reads (fixes thundering herd for massive batch calculations)
  if (pendingGets.has(key)) return pendingGets.get(key)!;

  const promise = (async () => {
    const db = await initDB();
    return new Promise<string | null>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readonly");
      const store = tx.objectStore(STORE_NAME);
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result !== undefined ? req.result : null);
      req.onerror = () => reject(req.error);
    });
  })();

  pendingGets.set(key, promise);
  promise.finally(() => setTimeout(() => pendingGets.delete(key), 10)); // Clear after resolution

  return promise;
}

export async function idbSet(key: string, value: string): Promise<void> {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    const store = tx.objectStore(STORE_NAME);
    const req = store.put(value, key);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}