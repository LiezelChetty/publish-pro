import { strFromU8 } from "fflate";
import { getAttr, getElementsByLocalName, getXml, normalizeRelationshipTarget, type DocxPackage } from "./docxPackage";
import type { DocxExtractedImage, DocxImportWarning } from "./types";

export type DocxRelationships = {
  targets: Map<string, { type: string; target: string; external: boolean }>;
};

export function parseDocumentRelationships(pkg: DocxPackage): DocxRelationships {
  const xml = getXml(pkg, "word/_rels/document.xml.rels", "word/_rels/document.xml.rels");
  const targets = new Map<string, { type: string; target: string; external: boolean }>();
  if (!xml) return { targets };
  for (const relationship of getElementsByLocalName(xml, "Relationship")) {
    const id = getAttr(relationship, "Id");
    const type = getAttr(relationship, "Type") ?? "";
    const target = getAttr(relationship, "Target") ?? "";
    if (!id || !target) continue;
    targets.set(id, {
      type,
      target: type.includes("hyperlink") ? target : normalizeRelationshipTarget(target),
      external: getAttr(relationship, "TargetMode") === "External",
    });
  }
  return { targets };
}

export async function extractDocxImages(pkg: DocxPackage, relationships: DocxRelationships, warnings: DocxImportWarning[]) {
  const images: DocxExtractedImage[] = [];
  const seenHashes = new Map<string, DocxExtractedImage>();
  for (const [relationshipId, relationship] of relationships.targets.entries()) {
    if (!relationship.type.includes("/image")) continue;
    const bytes = pkg.files[relationship.target];
    if (!bytes) {
      warnings.push({ code: "missing-image", message: `Image relationship ${relationshipId} points to a missing file.` });
      continue;
    }
    const mimeType = getImageMimeType(relationship.target);
    if (mimeType === "image/unsupported") {
      warnings.push({ code: "unsupported-image", message: `${relationship.target} is not a supported first-phase image format.` });
      continue;
    }
    const hash = quickContentHash(bytes);
    const existing = seenHashes.get(hash);
    if (existing) {
      images.push({ ...existing, relationshipId });
      continue;
    }
    const name = relationship.target.split("/").pop() ?? "Image";
    const dataUrl = `data:${mimeType};base64,${bytesToBase64(bytes)}`;
    const size = getImageSize(bytes, mimeType);
    const image: DocxExtractedImage = {
      id: crypto.randomUUID(),
      relationshipId,
      name,
      path: relationship.target,
      mimeType,
      bytes,
      dataUrl,
      width: size.width,
      height: size.height,
      used: false,
    };
    seenHashes.set(hash, image);
    images.push(image);
  }
  return images;
}

export function getHyperlinkTarget(relationships: DocxRelationships, relationshipId?: string) {
  if (!relationshipId) return undefined;
  const relationship = relationships.targets.get(relationshipId);
  return relationship?.type.includes("hyperlink") ? relationship.target : undefined;
}

function getImageMimeType(path: string): DocxExtractedImage["mimeType"] {
  const lower = path.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".gif")) return "image/gif";
  if (lower.endsWith(".svg")) return "image/svg+xml";
  return "image/unsupported";
}

function getImageSize(bytes: Uint8Array, mimeType: DocxExtractedImage["mimeType"]) {
  if (mimeType === "image/png" && bytes.length > 24) {
    return {
      width: readUint32(bytes, 16),
      height: readUint32(bytes, 20),
    };
  }
  if (mimeType === "image/jpeg") return readJpegSize(bytes);
  if (mimeType === "image/svg+xml") {
    const text = strFromU8(bytes);
    const width = Number(text.match(/\bwidth=["']?([\d.]+)/i)?.[1]) || 320;
    const height = Number(text.match(/\bheight=["']?([\d.]+)/i)?.[1]) || 180;
    return { width, height };
  }
  return { width: 320, height: 180 };
}

function readJpegSize(bytes: Uint8Array) {
  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) break;
    const marker = bytes[offset + 1];
    const length = (bytes[offset + 2] << 8) + bytes[offset + 3];
    if (marker >= 0xc0 && marker <= 0xc3) {
      return { height: (bytes[offset + 5] << 8) + bytes[offset + 6], width: (bytes[offset + 7] << 8) + bytes[offset + 8] };
    }
    offset += 2 + length;
  }
  return { width: 320, height: 180 };
}

function readUint32(bytes: Uint8Array, offset: number) {
  return ((bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3]) >>> 0;
}

function bytesToBase64(bytes: Uint8Array) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
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
