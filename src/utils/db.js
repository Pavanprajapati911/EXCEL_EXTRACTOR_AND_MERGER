const DB_NAME = 'ExcelMergerDB';
const STORE_NAME = 'workspace';
const KEY_NAME = 'currentState';

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);

    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };

    request.onsuccess = (e) => {
      resolve(e.target.result);
    };

    request.onerror = (e) => {
      reject(new Error(`Failed to open IndexedDB: ${e.target.error?.message}`));
    };
  });
}

/**
 * Saves the current workspace state to IndexedDB
 */
export async function saveWorkspace(files, rowConfigs, columnMappings, outputFileName, sheetLocations) {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      
      const stateObj = {
        files,
        rowConfigs,
        columnMappings,
        outputFileName,
        sheetLocations,
        updatedAt: Date.now()
      };

      const request = store.put(stateObj, KEY_NAME);

      request.onsuccess = () => resolve(true);
      request.onerror = (e) => reject(new Error(`Failed to save state: ${e.target.error?.message}`));
    });
  } catch (error) {
    console.error("IndexedDB Save Error:", error);
    return false;
  }
}

/**
 * Loads the saved workspace state from IndexedDB
 * @returns {Promise<Object|null>} Saved state object or null if none exists
 */
export async function loadWorkspace() {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readonly');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get(KEY_NAME);

      request.onsuccess = (e) => {
        resolve(e.target.result || null);
      };
      request.onerror = (e) => {
        reject(new Error(`Failed to load state: ${e.target.error?.message}`));
      };
    });
  } catch (error) {
    console.error("IndexedDB Load Error:", error);
    return null;
  }
}

/**
 * Clears the saved workspace state from IndexedDB
 */
export async function clearWorkspace() {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.delete(KEY_NAME);

      request.onsuccess = () => resolve(true);
      request.onerror = (e) => reject(new Error(`Failed to clear state: ${e.target.error?.message}`));
    });
  } catch (error) {
    console.error("IndexedDB Clear Error:", error);
    return false;
  }
}
