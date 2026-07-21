import type { ProjectManifest, ProjectMetadata } from "./schema";

const AUTOSAVE_KEY = "publish-pro-autosave-v1";
const RECENT_PROJECTS_KEY = "publish-pro-recent-projects-v1";
const MAX_RECENT_PROJECTS = 8;
const MAX_RECENT_BYTES = 4_500_000;

export type ProjectAutosave = {
  savedAt: string;
  dataBase64: string;
  manifest: ProjectManifest;
};

export type RecentProject = {
  id: string;
  name: string;
  lastOpenedAt: string;
  pageCount: number;
  sourceName?: string;
  dataBase64?: string;
};

export function saveAutosave(bytes: Uint8Array, manifest: ProjectManifest) {
  if (bytes.byteLength > MAX_RECENT_BYTES) return false;
  const payload: ProjectAutosave = {
    savedAt: new Date().toISOString(),
    dataBase64: uint8ToBase64(bytes),
    manifest,
  };
  localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(payload));
  return true;
}

export function loadAutosave(): ProjectAutosave | null {
  return parseStorageItem<ProjectAutosave>(AUTOSAVE_KEY);
}

export function clearAutosave() {
  localStorage.removeItem(AUTOSAVE_KEY);
}

export function loadRecentProjects(): RecentProject[] {
  return parseStorageItem<RecentProject[]>(RECENT_PROJECTS_KEY) ?? [];
}

export function addRecentProject(bytes: Uint8Array, manifest: ProjectManifest) {
  const existing = loadRecentProjects().filter((project) => project.id !== manifest.projectId);
  const sourceName = manifest.sources[0]?.name;
  const next: RecentProject = {
    id: manifest.projectId,
    name: manifest.metadata.name,
    lastOpenedAt: new Date().toISOString(),
    pageCount: manifest.pages.length,
    sourceName,
    dataBase64: bytes.byteLength <= MAX_RECENT_BYTES ? uint8ToBase64(bytes) : undefined,
  };
  localStorage.setItem(RECENT_PROJECTS_KEY, JSON.stringify([next, ...existing].slice(0, MAX_RECENT_PROJECTS)));
}

export function removeRecentProject(id: string) {
  localStorage.setItem(RECENT_PROJECTS_KEY, JSON.stringify(loadRecentProjects().filter((project) => project.id !== id)));
}

export function uint8ToBase64(bytes: Uint8Array) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.slice(index, index + chunkSize));
  }
  return btoa(binary);
}

export function base64ToUint8(value: string) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
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

function parseStorageItem<T>(key: string): T | null {
  try {
    const raw = localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : null;
  } catch {
    return null;
  }
}
