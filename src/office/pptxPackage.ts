import { unzipSync } from "fflate";
import { getAttr, getElementsByLocalName, parseXml } from "./docxPackage";
import type { PptxPackage, PptxRelationship } from "./pptxTypes";

const decoder = new TextDecoder();

export function isPptxFile(file: File) {
  return file.name.toLowerCase().endsWith(".pptx") || file.type === "application/vnd.openxmlformats-officedocument.presentationml.presentation";
}

export function rejectUnsupportedPowerPointFile(file: File) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".ppt") && !name.endsWith(".pptx")) throw new Error("Legacy .ppt files are not supported yet. Choose a .pptx PowerPoint presentation.");
  if (!isPptxFile(file)) throw new Error("Choose a .pptx PowerPoint presentation.");
}

export function readPptxPackage(bytes: Uint8Array): PptxPackage {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch (error) {
    throw new Error(error instanceof Error ? `The PPTX package could not be opened. It may be corrupt or password-protected. ${error.message}` : "The PPTX package could not be opened.");
  }

  if (!files["[Content_Types].xml"]) throw new Error("This file is missing [Content_Types].xml and is not a valid PPTX package.");
  if (!files["ppt/presentation.xml"]) throw new Error("This PPTX is missing ppt/presentation.xml.");
  if (!files["ppt/_rels/presentation.xml.rels"]) throw new Error("This PPTX is missing presentation relationships.");
  if (files["EncryptionInfo"] || files["EncryptedPackage"]) throw new Error("Password-protected or encrypted PPTX files are not supported yet.");

  return { files, textFiles: new Map(), warnings: [] };
}

export function getPptxPackageText(pkg: PptxPackage, path: string) {
  if (pkg.textFiles.has(path)) return pkg.textFiles.get(path) ?? "";
  const bytes = pkg.files[path];
  if (!bytes) return "";
  const text = decoder.decode(bytes);
  pkg.textFiles.set(path, text);
  return text;
}

export function getPptxXml(pkg: PptxPackage, path: string, description = path) {
  const text = getPptxPackageText(pkg, path);
  if (!text) return null;
  return parseXml(text, description);
}

export function parsePptxRelationships(pkg: PptxPackage, relsPath: string, basePath = "ppt") {
  const xml = getPptxXml(pkg, relsPath, relsPath);
  const relationships = new Map<string, PptxRelationship>();
  if (!xml) return relationships;
  for (const relationship of getElementsByLocalName(xml, "Relationship")) {
    const id = getAttr(relationship, "Id");
    const type = getAttr(relationship, "Type") ?? "";
    const rawTarget = getAttr(relationship, "Target") ?? "";
    if (!id || !rawTarget) continue;
    const external = getAttr(relationship, "TargetMode") === "External";
    relationships.set(id, {
      type,
      target: external ? rawTarget : normalizePptxTarget(rawTarget, basePath),
      external,
    });
  }
  return relationships;
}

export function normalizePptxTarget(target: string, basePath = "ppt") {
  if (target.startsWith("/")) return target.slice(1);
  if (target.startsWith("ppt/")) return target;
  const baseParts = basePath.split("/").filter(Boolean);
  const targetParts = target.split("/").filter(Boolean);
  for (const part of targetParts) {
    if (part === ".") continue;
    if (part === "..") baseParts.pop();
    else baseParts.push(part);
  }
  return baseParts.join("/");
}
