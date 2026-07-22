import { BaseDirectory, appDataDir, basename } from "@tauri-apps/api/path";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { listen } from "@tauri-apps/api/event";
import { open, save } from "@tauri-apps/plugin-dialog";
import { exists, mkdir, readFile, readTextFile, writeFile, writeTextFile } from "@tauri-apps/plugin-fs";
import { openPath as openSystemPath, revealItemInDir } from "@tauri-apps/plugin-opener";
import type { ProjectManifest } from "../projects/schema";
import type { ProjectAutosave, RecentProject } from "../projects/storage";
import type { RuntimeAdapter, RuntimeFile, RuntimeFileKind, RuntimeOpenOptions, RuntimeSaveOptions } from "./types";

const DESKTOP_RECENTS_LIMIT = 12;
const AUTOSAVE_PATH = "autosaves/latest.pproj";
const AUTOSAVE_INDEX_PATH = "autosaves/index.json";
const RECENTS_PATH = "recent-projects.json";

export function isTauriRuntime() {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

export const desktopRuntime: RuntimeAdapter = {
  isDesktop: true,
  async openFiles(options) {
    const selected = await open({
      title: options.title ?? getOpenTitle(options.kind),
      multiple: options.multiple,
      filters: getOpenFilters(options.kind),
      fileAccessMode: "scoped",
    });
    const paths = Array.isArray(selected) ? selected : selected ? [selected] : [];
    if (paths.length === 0) return [];
    const files: RuntimeFile[] = [];
    for (const path of paths) {
      const bytes = await readFile(path);
      const name = await basename(path);
      files.push({ name, path, bytes, mimeType: getMimeType(name) });
    }
    return files;
  },
  async readPath(path) {
    const bytes = await readFile(path);
    const name = await basename(path);
    return { name, path, bytes, mimeType: getMimeType(name) };
  },
  async saveFile(options) {
    const targetPath = options.currentPath ?? await save({
      title: options.kind === "project" ? "Save Publish Pro Project" : "Publish PDF",
      defaultPath: options.suggestedName,
      filters: options.kind === "project" ? projectFilters : pdfFilters,
      canCreateDirectories: true,
    });
    if (!targetPath) return { saved: false, native: true };
    await writeFile(targetPath, options.bytes);
    return { saved: true, path: targetPath, native: true };
  },
  async setWindowTitle(title) {
    await getCurrentWindow().setTitle(title);
  },
  async revealPath(path) {
    await revealItemInDir(path);
  },
  async openPath(path) {
    await openSystemPath(path);
  },
  async forceClose() {
    await invoke("force_close");
  },
  async loadAutosave() {
    try {
      const hasAutosave = await exists(AUTOSAVE_PATH, { baseDir: BaseDirectory.AppData });
      const hasIndex = await exists(AUTOSAVE_INDEX_PATH, { baseDir: BaseDirectory.AppData });
      if (!hasAutosave || !hasIndex) return null;
      const [bytes, indexRaw] = await Promise.all([
        readFile(AUTOSAVE_PATH, { baseDir: BaseDirectory.AppData }),
        readTextFile(AUTOSAVE_INDEX_PATH, { baseDir: BaseDirectory.AppData }),
      ]);
      const index = JSON.parse(indexRaw) as { savedAt: string; manifest: ProjectManifest };
      return {
        id: "latest",
        savedAt: index.savedAt,
        bytes,
        manifest: index.manifest,
      };
    } catch {
      return null;
    }
  },
  async saveAutosave(bytes, manifest, sourcePath) {
    await ensureAppDataDir("autosaves");
    const savedAt = new Date().toISOString();
    const payload: ProjectAutosave = { id: "latest", savedAt, bytes, manifest };
    await writeFile(AUTOSAVE_PATH, bytes, { baseDir: BaseDirectory.AppData });
    await writeTextFile(AUTOSAVE_INDEX_PATH, JSON.stringify({ version: 1, savedAt, sourcePath, manifest }, null, 2), { baseDir: BaseDirectory.AppData });
    return payload;
  },
  async clearAutosave() {
    await ensureAppDataDir("autosaves");
    await writeTextFile(AUTOSAVE_INDEX_PATH, JSON.stringify({ version: 1, clearedAt: new Date().toISOString() }, null, 2), { baseDir: BaseDirectory.AppData });
  },
  async loadRecentProjects() {
    try {
      if (!(await exists(RECENTS_PATH, { baseDir: BaseDirectory.AppData }))) return [];
      const raw = await readTextFile(RECENTS_PATH, { baseDir: BaseDirectory.AppData });
      return JSON.parse(raw) as RecentProject[];
    } catch {
      return [];
    }
  },
  async addRecentProject(manifest, options = {}) {
    const existing = await this.loadRecentProjects() ?? [];
    const id = options.path ?? manifest.projectId;
    const next: RecentProject = {
      id,
      name: manifest.metadata.name,
      lastOpenedAt: new Date().toISOString(),
      lastModifiedAt: manifest.metadata.modifiedAt,
      pageCount: manifest.pages.length,
      sourceName: manifest.sources[0]?.name,
      savedFilename: options.filename ?? (options.path ? await basename(options.path) : undefined),
      origin: options.origin ?? "file-system-access",
      autosaveAvailable: options.autosaveAvailable,
    };
    const deduped = existing.filter((project) => project.id !== id && project.savedFilename !== next.savedFilename);
    await ensureAppDataDir("");
    await writeTextFile(RECENTS_PATH, JSON.stringify([next, ...deduped].slice(0, DESKTOP_RECENTS_LIMIT), null, 2), { baseDir: BaseDirectory.AppData });
  },
  async removeRecentProject(id) {
    const existing = await this.loadRecentProjects() ?? [];
    await ensureAppDataDir("");
    await writeTextFile(RECENTS_PATH, JSON.stringify(existing.filter((project) => project.id !== id), null, 2), { baseDir: BaseDirectory.AppData });
  },
  async clearRecentProjects() {
    await ensureAppDataDir("");
    await writeTextFile(RECENTS_PATH, "[]", { baseDir: BaseDirectory.AppData });
  },
  async listenMenuActions(handler) {
    return listen<string>("desktop://menu-action", (event) => handler(event.payload));
  },
  async listenLaunchFiles(handler) {
    return listen<string[]>("desktop://launch-files", (event) => handler(event.payload));
  },
};

const projectFilters = [{ name: "Publish Pro Project", extensions: ["pproj"] }];
const pdfFilters = [{ name: "PDF Document", extensions: ["pdf"] }];
const officeFilters = [
  { name: "Supported Documents", extensions: ["pproj", "pdf", "docx", "pptx"] },
  ...projectFilters,
  ...pdfFilters,
  { name: "Word Document", extensions: ["docx"] },
  { name: "PowerPoint Presentation", extensions: ["pptx"] },
];

function getOpenFilters(kind: RuntimeFileKind) {
  if (kind === "project") return projectFilters;
  if (kind === "pdf") return pdfFilters;
  if (kind === "docx") return [{ name: "Word Document", extensions: ["docx"] }];
  if (kind === "pptx") return [{ name: "PowerPoint Presentation", extensions: ["pptx"] }];
  if (kind === "image") return [{ name: "Images", extensions: ["png", "jpg", "jpeg", "svg", "webp"] }];
  return officeFilters;
}

function getOpenTitle(kind: RuntimeFileKind) {
  if (kind === "project") return "Open Publish Pro Project";
  if (kind === "pdf") return "Open PDF";
  if (kind === "docx") return "Import Word Document";
  if (kind === "pptx") return "Import PowerPoint Presentation";
  return "Open File";
}

function getMimeType(name: string) {
  const lower = name.toLowerCase();
  if (lower.endsWith(".pproj")) return "application/vnd.publish-pro.project+zip";
  if (lower.endsWith(".pdf")) return "application/pdf";
  if (lower.endsWith(".docx")) return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (lower.endsWith(".pptx")) return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".svg")) return "image/svg+xml";
  if (lower.endsWith(".webp")) return "image/webp";
  return "application/octet-stream";
}

async function ensureAppDataDir(path: string) {
  if (path) {
    await mkdir(path, { baseDir: BaseDirectory.AppData, recursive: true });
    return;
  }
  await mkdir(await appDataDir(), { recursive: true });
}

export async function runtimeFileToBrowserFile(file: RuntimeFile) {
  return new File([file.bytes.slice()], file.name, { type: file.mimeType });
}
