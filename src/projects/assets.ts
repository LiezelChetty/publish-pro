import type { ProjectAssetManifest } from "./schema";

type SourceLike = {
  id: string;
  name: string;
  bytes: Uint8Array;
};

type PageLike = {
  sourceDocumentId?: string;
};

type MarkLike = {
  id: string;
  kind: string;
  imageDataUrl?: string;
  imageMimeType?: string;
  imageName?: string;
};

export type ProjectAssetSummary = ProjectAssetManifest & {
  firstMarkId?: string;
  firstPageIndex?: number;
};

export function collectProjectAssets(sources: Record<string, SourceLike>, pages: PageLike[], marks: MarkLike[]): ProjectAssetSummary[] {
  const collectedAt = new Date().toISOString();
  const sourceUsage = new Map<string, number>();
  pages.forEach((page) => {
    if (page.sourceDocumentId) sourceUsage.set(page.sourceDocumentId, (sourceUsage.get(page.sourceDocumentId) ?? 0) + 1);
  });

  const sourceAssets = Object.values(sources).map((source) => ({
    id: source.id,
    name: source.name,
    type: "source-pdf" as const,
    path: `sources/${source.id}.pdf`,
    mimeType: "application/pdf",
    size: source.bytes.byteLength,
    dateAdded: collectedAt,
    contentHash: quickContentHash(source.bytes),
    usageCount: sourceUsage.get(source.id) ?? 0,
    status: "available" as const,
    firstPageIndex: pages.findIndex((page) => page.sourceDocumentId === source.id),
  }));

  const embeddedAssets = marks
    .filter((mark) => (mark.kind === "image" || mark.kind === "pngSignature") && mark.imageDataUrl)
    .map((mark) => ({
      id: mark.id,
      name: mark.imageName || (mark.kind === "pngSignature" ? "Signature.png" : "Image"),
      type: mark.kind === "pngSignature" ? ("signature" as const) : ("image" as const),
      path: `assets/${mark.id}`,
      mimeType: mark.imageMimeType,
      size: estimateDataUrlSize(mark.imageDataUrl ?? ""),
      dateAdded: collectedAt,
      contentHash: mark.imageDataUrl ? quickStringHash(mark.imageDataUrl) : undefined,
      usageCount: 1,
      status: "available" as const,
      firstMarkId: mark.id,
    }));

  return [...sourceAssets, ...embeddedAssets];
}

function quickContentHash(bytes: Uint8Array) {
  let hash = 2166136261;
  const step = Math.max(1, Math.floor(bytes.byteLength / 4096));
  for (let index = 0; index < bytes.byteLength; index += step) {
    hash ^= bytes[index];
    hash = Math.imul(hash, 16777619);
  }
  return `fnv1a-${(hash >>> 0).toString(16)}-${bytes.byteLength}`;
}

function quickStringHash(value: string) {
  let hash = 2166136261;
  const step = Math.max(1, Math.floor(value.length / 4096));
  for (let index = 0; index < value.length; index += step) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return `fnv1a-${(hash >>> 0).toString(16)}-${value.length}`;
}

export function estimateDataUrlSize(dataUrl: string) {
  const encoded = dataUrl.split(",", 2)[1] ?? "";
  return Math.floor((encoded.length * 3) / 4);
}

export function formatAssetSize(size: number) {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${Math.round(size / 102.4) / 10} KB`;
  return `${Math.round(size / 1024 / 102.4) / 10} MB`;
}
