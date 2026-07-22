import type { ProjectManifest, ProjectMetadata } from "./schema";

const DB_NAME = "publish-pro-projects";
const DB_VERSION = 1;
const AUTOSAVE_STORE = "autosave";
const AUTOSAVE_ID = "latest";
const RECENT_PROJECTS_KEY = "publish-pro-recent-projects-v2";
const MAX_RECENT_PROJECTS = 12;

export type ProjectAutosave = {
  id: typeof AUTOSAVE_ID;
  savedAt: string;
  bytes: Uint8Array;
  manifest: ProjectManifest;
};

export type RecentProject = {
  id: string;
  name: string;
  lastOpenedAt: string;
  lastModifiedAt: string;
  pageCount: number;
  sourceName?: string;
  savedFilename?: string;
  origin: "download" | "file-system-access" | "browser-import" | "autosave";
  autosaveAvailable?: boolean;
};

export async function saveAutosave(bytes: Uint8Array, manifest: ProjectManifest) {
  const db = await openProjectDatabase();
  const payload: ProjectAutosave = {
    id: AUTOSAVE_ID,
    savedAt: new Date().toISOString(),
    bytes,
    manifest,
  };
  await putValue(db, AUTOSAVE_STORE, payload);
  return payload;
}

export async function loadAutosave(): Promise<ProjectAutosave | null> {
  const db = await openProjectDatabase();
  return (await getValue<ProjectAutosave>(db, AUTOSAVE_STORE, AUTOSAVE_ID)) ?? null;
}

export async function clearAutosave() {
  const db = await openProjectDatabase();
  await deleteValue(db, AUTOSAVE_STORE, AUTOSAVE_ID);
}

export function loadRecentProjects(): RecentProject[] {
  return parseStorageItem<RecentProject[]>(RECENT_PROJECTS_KEY) ?? [];
}

export async function addRecentProject(manifest: ProjectManifest, options: { filename?: string; origin?: RecentProject["origin"]; autosaveAvailable?: boolean } = {}) {
  const existing = loadRecentProjects().filter((project) => project.id !== manifest.projectId);
  const autosave = options.autosaveAvailable ?? Boolean(await loadAutosave());
  const next: RecentProject = {
    id: manifest.projectId,
    name: manifest.metadata.name,
    lastOpenedAt: new Date().toISOString(),
    lastModifiedAt: manifest.metadata.modifiedAt,
    pageCount: manifest.pages.length,
    sourceName: manifest.sources[0]?.name,
    savedFilename: options.filename,
    origin: options.origin ?? "download",
    autosaveAvailable: autosave,
  };
  localStorage.setItem(RECENT_PROJECTS_KEY, JSON.stringify([next, ...existing].slice(0, MAX_RECENT_PROJECTS)));
}

export function removeRecentProject(id: string) {
  localStorage.setItem(RECENT_PROJECTS_KEY, JSON.stringify(loadRecentProjects().filter((project) => project.id !== id)));
}

export function clearRecentProjects() {
  localStorage.removeItem(RECENT_PROJECTS_KEY);
}

export function metadataTagsToInput(metadata: ProjectMetadata) {
  return metadata.tags.join(", ");
}

export function inputToMetadataTags(value: string) {
  return value
    .split(",")
    .map((tag) => tag.trim())
    .filter(Boolean);
}

function openProjectDatabase() {
  return new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(AUTOSAVE_STORE)) db.createObjectStore(AUTOSAVE_STORE, { keyPath: "id" });
    };
    request.onerror = () => reject(request.error ?? new Error("Could not open Publish Pro project storage."));
    request.onsuccess = () => resolve(request.result);
  });
}

function putValue<T>(db: IDBDatabase, storeName: string, value: T) {
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(storeName, "readwrite");
    transaction.objectStore(storeName).put(value);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error("Could not write project storage."));
  });
}

function getValue<T>(db: IDBDatabase, storeName: string, key: string) {
  return new Promise<T | undefined>((resolve, reject) => {
    const transaction = db.transaction(storeName, "readonly");
    const request = transaction.objectStore(storeName).get(key);
    request.onsuccess = () => resolve(request.result as T | undefined);
    request.onerror = () => reject(request.error ?? new Error("Could not read project storage."));
  });
}

function deleteValue(db: IDBDatabase, storeName: string, key: string) {
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(storeName, "readwrite");
    transaction.objectStore(storeName).delete(key);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error("Could not clear project storage."));
  });
}

function parseStorageItem<T>(key: string): T | null {
  try {
    const raw = localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : null;
  } catch {
    return null;
  }
}
