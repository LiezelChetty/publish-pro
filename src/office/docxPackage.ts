import { unzipSync } from "fflate";
import type { DocxImportWarning } from "./types";

const decoder = new TextDecoder();

export type DocxPackage = {
  files: Record<string, Uint8Array>;
  textFiles: Map<string, string>;
  warnings: DocxImportWarning[];
};

export function isDocxFile(file: File) {
  return file.name.toLowerCase().endsWith(".docx") || file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
}

export function rejectUnsupportedWordFile(file: File) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".doc") && !name.endsWith(".docx")) {
    throw new Error("Legacy .doc files are not supported yet. Choose a .docx Word document.");
  }
  if (!isDocxFile(file)) {
    throw new Error("Choose a .docx Word document. Legacy .doc files are not supported in this phase.");
  }
}

export function readDocxPackage(bytes: Uint8Array): DocxPackage {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch (error) {
    throw new Error(error instanceof Error ? `The DOCX package could not be opened. It may be corrupt or password-protected. ${error.message}` : "The DOCX package could not be opened.");
  }

  const warnings: DocxImportWarning[] = [];
  if (!files["[Content_Types].xml"]) throw new Error("This file is missing [Content_Types].xml and is not a valid DOCX package.");
  if (!files["word/document.xml"]) throw new Error("This DOCX is missing word/document.xml.");
  if (files["EncryptionInfo"] || files["EncryptedPackage"]) {
    throw new Error("Password-protected or encrypted DOCX files are not supported yet.");
  }
  if (!files["word/_rels/document.xml.rels"]) {
    warnings.push({ code: "missing-relationships", message: "The DOCX is missing document relationships; images and links may not import." });
  }

  return { files, textFiles: new Map(), warnings };
}

export function getPackageText(pkg: DocxPackage, path: string) {
  if (pkg.textFiles.has(path)) return pkg.textFiles.get(path) ?? "";
  const bytes = pkg.files[path];
  if (!bytes) return "";
  const text = decoder.decode(bytes);
  pkg.textFiles.set(path, text);
  return text;
}

export function parseXml(text: string, description: string) {
  const parser = new DOMParser();
  const xml = parser.parseFromString(text, "application/xml");
  const parserError = xml.getElementsByTagName("parsererror")[0];
  if (parserError) throw new Error(`${description} contains invalid XML.`);
  return xml;
}

export function getXml(pkg: DocxPackage, path: string, description: string) {
  const text = getPackageText(pkg, path);
  if (!text) return null;
  return parseXml(text, description);
}

export function getElementsByLocalName(parent: Element | Document, localName: string): Element[] {
  return Array.from(parent.getElementsByTagName("*")).filter((element): element is Element => element.localName === localName);
}

export function getFirstChildByLocalName(parent: ParentNode, localName: string) {
  return Array.from(parent.childNodes).find((node): node is Element => node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === localName);
}

export function getChildrenByLocalName(parent: ParentNode, localName: string) {
  return Array.from(parent.childNodes).filter((node): node is Element => node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === localName);
}

export function getAttr(element: Element | null | undefined, localName: string) {
  if (!element) return undefined;
  for (const attr of Array.from(element.attributes)) {
    if (attr.localName === localName) return attr.value;
  }
  return undefined;
}

export function getText(element: Element | Document | null | undefined) {
  if (!element) return "";
  return Array.from(element.getElementsByTagName("*"))
    .filter((item) => item.localName === "t")
    .map((item) => item.textContent ?? "")
    .join("");
}

export function normalizeRelationshipTarget(target: string) {
  if (target.startsWith("/")) return target.slice(1);
  if (target.startsWith("word/")) return target;
  return `word/${target.replace(/^\.\//, "")}`;
}
