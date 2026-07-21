export const PROJECT_FORMAT = "publish-pro-project";
export const PROJECT_FORMAT_VERSION = 1;

export type ProjectAssetType = "source-pdf" | "image" | "signature" | "other";
export type ProjectAssetStatus = "available" | "missing" | "invalid";

export type ProjectMetadata = {
  name: string;
  client: string;
  author: string;
  description: string;
  tags: string[];
  createdAt: string;
  modifiedAt: string;
};

export type ProjectAssetManifest = {
  id: string;
  name: string;
  type: ProjectAssetType;
  path: string;
  mimeType?: string;
  size: number;
  usageCount: number;
  status: ProjectAssetStatus;
};

export type ProjectWorkspaceState = {
  currentPage: number;
  zoom: number;
  activeWorkspace: string;
  leftPanel: string;
  selectedPageIds: string[];
};

export type ProjectExportSettings = {
  defaultFileName: string;
};

export type ProjectManifest = {
  format: typeof PROJECT_FORMAT;
  version: typeof PROJECT_FORMAT_VERSION;
  projectId: string;
  metadata: ProjectMetadata;
  workspaceState: ProjectWorkspaceState;
  exportSettings: ProjectExportSettings;
  pages: unknown[];
  annotations: unknown[];
  sources: ProjectAssetManifest[];
  assets: ProjectAssetManifest[];
};

export type ProjectBundle = {
  manifest: ProjectManifest;
  files: Record<string, Uint8Array>;
};

export function createProjectMetadata(name = "Untitled project"): ProjectMetadata {
  const now = new Date().toISOString();
  return {
    name,
    client: "",
    author: "",
    description: "",
    tags: [],
    createdAt: now,
    modifiedAt: now,
  };
}

export function assertSupportedManifest(value: unknown): asserts value is ProjectManifest {
  if (!value || typeof value !== "object") {
    throw new Error("Project manifest is missing or invalid.");
  }

  const manifest = value as Partial<ProjectManifest>;
  if (manifest.format !== PROJECT_FORMAT) {
    throw new Error("This is not a Publish Pro project file.");
  }
  if (typeof manifest.version !== "number") {
    throw new Error("Project file version is missing.");
  }
  if (manifest.version > PROJECT_FORMAT_VERSION) {
    throw new Error(`This project was created by a newer Publish Pro format version (${manifest.version}).`);
  }
  if (!Array.isArray(manifest.pages) || !Array.isArray(manifest.annotations)) {
    throw new Error("Project manifest is missing document data.");
  }
  if (!manifest.metadata || !manifest.workspaceState || !manifest.exportSettings) {
    throw new Error("Project manifest is incomplete.");
  }
}
