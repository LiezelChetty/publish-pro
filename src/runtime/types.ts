import type { ProjectAutosave, RecentProject } from "../projects/storage";
import type { ProjectManifest } from "../projects/schema";

export type RuntimeFileKind = "project" | "pdf" | "docx" | "pptx" | "image" | "any";

export type RuntimeFile = {
  name: string;
  path?: string;
  mimeType: string;
  bytes: Uint8Array;
};

export type RuntimeSaveResult = {
  saved: boolean;
  path?: string;
  native: boolean;
};

export type RuntimeOpenOptions = {
  kind: RuntimeFileKind;
  multiple?: boolean;
  title?: string;
};

export type RuntimeSaveOptions = {
  kind: "project" | "pdf";
  bytes: Uint8Array;
  suggestedName: string;
  currentPath?: string | null;
};

export type RuntimeAdapter = {
  isDesktop: boolean;
  openFiles(options: RuntimeOpenOptions): Promise<RuntimeFile[] | null>;
  readPath(path: string): Promise<RuntimeFile | null>;
  saveFile(options: RuntimeSaveOptions): Promise<RuntimeSaveResult | null>;
  setWindowTitle(title: string): Promise<void>;
  revealPath(path: string): Promise<void>;
  openPath(path: string): Promise<void>;
  forceClose(): Promise<void>;
  loadAutosave(): Promise<ProjectAutosave | null>;
  saveAutosave(bytes: Uint8Array, manifest: ProjectManifest, sourcePath?: string | null): Promise<ProjectAutosave | null>;
  clearAutosave(): Promise<void>;
  loadRecentProjects(): Promise<RecentProject[] | null>;
  addRecentProject(manifest: ProjectManifest, options?: { filename?: string; path?: string; origin?: RecentProject["origin"]; autosaveAvailable?: boolean }): Promise<void>;
  removeRecentProject(id: string): Promise<void>;
  clearRecentProjects(): Promise<void>;
  listenMenuActions(handler: (action: string) => void): Promise<() => void>;
  listenLaunchFiles(handler: (paths: string[]) => void): Promise<() => void>;
};
