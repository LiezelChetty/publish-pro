import { collectProjectAssets, type ProjectImageAssetLike } from "./assets";
import { createDefaultPublishingSettings } from "../publishing/defaults";
import type { PublishingSettings } from "../publishing/types";
import type { NavigationState } from "../navigation/types";
import { migrateProjectManifest } from "./migrations";
import { assertSupportedManifest, createProjectMetadata, PROJECT_FORMAT, PROJECT_FORMAT_VERSION, type ProjectBundle, type ProjectManifest, type ProjectMetadata } from "./schema";
import { createZip, readZip } from "./zip";

const decoder = new TextDecoder();

type SourceLike = {
  id: string;
  name: string;
  bytes: Uint8Array;
  mimeType?: string;
};

type CreateProjectInput = {
  projectId: string;
  metadata: ProjectMetadata;
  pages: unknown[];
  annotations: unknown[];
  projectImageAssets?: ProjectImageAssetLike[];
  publishingSettings?: PublishingSettings;
  navigation?: NavigationState;
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
    input.annotations as Array<{ id: string; kind: string; imageDataUrl?: string; imageMimeType?: string; imageName?: string }>,
    input.projectImageAssets,
    collectPublishingImageAssetIds(input.publishingSettings)
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
    publishingSettings: input.publishingSettings ?? createDefaultPublishingSettings(),
    navigation: input.navigation,
    sources: allAssets.filter((asset) => asset.type === "source-pdf" || asset.type === "source-docx"),
    assets: allAssets.filter((asset) => asset.type !== "source-pdf" && asset.type !== "source-docx"),
  };
}

function collectPublishingImageAssetIds(settings?: PublishingSettings) {
  const ids: string[] = [];
  if (!settings) return ids;
  if (settings.watermark.type === "image" && settings.watermark.imageAssetId) ids.push(settings.watermark.imageAssetId);
  for (const area of ["header", "footer"] as const) {
    for (const side of ["left", "center", "right"] as const) {
      const assetId = settings.headerFooter[area][side].image?.assetId;
      if (assetId) ids.push(assetId);
    }
  }
  for (const key of ["firstPage", "oddPage", "evenPage"] as const) {
    const override = settings.headerFooter[key];
    if (!override) continue;
    for (const zone of Object.values(override)) {
      if (typeof zone === "object" && zone.image?.assetId) ids.push(zone.image.assetId);
    }
  }
  return ids;
}

export function serializeProject(input: CreateProjectInput) {
  const manifest = buildProjectManifest(input);
  const files: Record<string, Uint8Array | string> = {
    "manifest.json": JSON.stringify(manifest, null, 2),
  };
  const sourcePaths = new Map(manifest.sources.map((source) => [source.id, source.path]));
  for (const source of Object.values(input.sourceDocuments)) {
    files[sourcePaths.get(source.id) ?? `sources/${source.id}.pdf`] = source.bytes;
  }
  for (const mark of input.annotations as Array<{ id: string; kind: string; imageDataUrl?: string }>) {
    if ((mark.kind === "image" || mark.kind === "pngSignature") && mark.imageDataUrl) {
      files[`assets/${mark.id}`] = dataUrlToBytes(mark.imageDataUrl);
    }
  }
  for (const asset of input.projectImageAssets ?? []) {
    files[`assets/${asset.id}`] = dataUrlToBytes(asset.dataUrl);
  }
  return createZip(files);
}

export function deserializeProject(bytes: Uint8Array): ProjectBundle {
  const files = readZip(bytes);
  const manifestBytes = files["manifest.json"];
  if (!manifestBytes) throw new Error("Project file is missing manifest.json.");
  const parsed = JSON.parse(decoder.decode(manifestBytes)) as unknown;
  const manifest = migrateProjectManifest(parsed);
  validateProjectAssets(manifest, files);
  return { manifest, files };
}

function validateProjectAssets(manifest: ProjectManifest, files: Record<string, Uint8Array>) {
  const seen = new Set<string>();
  for (const source of manifest.sources) {
    if (seen.has(source.id)) throw new Error(`Project contains duplicate asset ID: ${source.id}`);
    seen.add(source.id);
    if (!files[source.path]) throw new Error(`Project is missing source asset: ${source.name}`);
  }
  for (const asset of manifest.assets) {
    if (seen.has(asset.id)) throw new Error(`Project contains duplicate asset ID: ${asset.id}`);
    seen.add(asset.id);
    if (asset.status === "available" && !files[asset.path]) throw new Error(`Project is missing embedded asset: ${asset.name}`);
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
