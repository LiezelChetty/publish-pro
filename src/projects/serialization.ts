import { collectProjectAssets } from "./assets";
import { assertSupportedManifest, createProjectMetadata, PROJECT_FORMAT, PROJECT_FORMAT_VERSION, type ProjectBundle, type ProjectManifest, type ProjectMetadata } from "./schema";
import { createZip, readZip } from "./zip";

const decoder = new TextDecoder();

type SourceLike = {
  id: string;
  name: string;
  bytes: Uint8Array;
};

type CreateProjectInput = {
  projectId: string;
  metadata: ProjectMetadata;
  pages: unknown[];
  annotations: unknown[];
  sourceDocuments: Record<string, SourceLike>;
  workspaceState: ProjectManifest["workspaceState"];
  exportSettings: ProjectManifest["exportSettings"];
};

export function buildProjectManifest(input: CreateProjectInput): ProjectManifest {
  const modifiedAt = new Date().toISOString();
  const metadata = {
    ...createProjectMetadata(input.metadata.name),
    ...input.metadata,
    modifiedAt,
  };
  const allAssets = collectProjectAssets(
    input.sourceDocuments,
    input.pages as Array<{ sourceDocumentId?: string }>,
    input.annotations as Array<{ id: string; kind: string; imageDataUrl?: string; imageMimeType?: string; imageName?: string }>
  );

  return {
    format: PROJECT_FORMAT,
    version: PROJECT_FORMAT_VERSION,
    projectId: input.projectId,
    metadata,
    workspaceState: input.workspaceState,
    exportSettings: input.exportSettings,
    pages: input.pages,
    annotations: input.annotations,
    sources: allAssets.filter((asset) => asset.type === "source-pdf"),
    assets: allAssets.filter((asset) => asset.type !== "source-pdf"),
  };
}

export function serializeProject(input: CreateProjectInput) {
  const manifest = buildProjectManifest(input);
  const files: Record<string, Uint8Array | string> = {
    "manifest.json": JSON.stringify(manifest, null, 2),
  };
  for (const source of Object.values(input.sourceDocuments)) {
    files[`sources/${source.id}.pdf`] = source.bytes;
  }
  for (const mark of input.annotations as Array<{ id: string; kind: string; imageDataUrl?: string }>) {
    if ((mark.kind === "image" || mark.kind === "pngSignature") && mark.imageDataUrl) {
      files[`assets/${mark.id}`] = dataUrlToBytes(mark.imageDataUrl);
    }
  }
  return createZip(files);
}

export function deserializeProject(bytes: Uint8Array): ProjectBundle {
  const files = readZip(bytes);
  const manifestBytes = files["manifest.json"];
  if (!manifestBytes) throw new Error("Project file is missing manifest.json.");
  const parsed = JSON.parse(decoder.decode(manifestBytes)) as unknown;
  assertSupportedManifest(parsed);
  validateProjectAssets(parsed, files);
  return { manifest: parsed, files };
}

function validateProjectAssets(manifest: ProjectManifest, files: Record<string, Uint8Array>) {
  const seen = new Set<string>();
  for (const source of manifest.sources) {
    if (seen.has(source.id)) throw new Error(`Project contains duplicate asset ID: ${source.id}`);
    seen.add(source.id);
    if (!files[source.path]) throw new Error(`Project is missing source asset: ${source.name}`);
  }
}

function dataUrlToBytes(dataUrl: string) {
  const [, encoded = ""] = dataUrl.split(",", 2);
  const binary = atob(encoded);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}
