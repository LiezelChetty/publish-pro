import { strFromU8 } from "fflate";
import { getAttr, getElementsByLocalName, getFirstChildByLocalName, getText } from "./docxPackage";
import { getPptxXml, parsePptxRelationships } from "./pptxPackage";
import type { PptxImageAsset, PptxImportOptions, PptxIntermediatePresentation, PptxPackage, PptxRelationship, PptxSlide, PptxSlideElement, PptxSlideSize, PptxTextRun } from "./pptxTypes";

export function parsePptxPresentation(pkg: PptxPackage, options: PptxImportOptions): PptxIntermediatePresentation {
  const presentationXml = getPptxXml(pkg, "ppt/presentation.xml", "ppt/presentation.xml");
  if (!presentationXml) throw new Error("The PPTX presentation XML is missing.");
  const presentationRels = parsePptxRelationships(pkg, "ppt/_rels/presentation.xml.rels", "ppt");
  const slideSize = parseSlideSize(presentationXml);
  const slideIdList = getElementsByLocalName(presentationXml, "sldId");
  const sections = parseSections(presentationXml);
  const imageAssets: PptxImageAsset[] = [];
  const imageByHash = new Map<string, PptxImageAsset>();
  const slides: PptxSlide[] = [];

  for (const [index, slideIdElement] of slideIdList.entries()) {
    const relId = getRelationshipAttr(slideIdElement, "id");
    const slideTarget = presentationRels.get(relId ?? "")?.target;
    if (!slideTarget) {
      pkg.warnings.push({ code: "pptx-missing-slide-relationship", category: "general", message: `Slide ${index + 1} does not have a resolvable presentation relationship and was skipped.` });
      continue;
    }
    const slide = parseSlide(pkg, slideTarget, index + 1, slideSize, imageAssets, imageByHash);
    if (slide.hidden && !options.includeHiddenSlides) continue;
    slides.push(slide);
  }

  if (slides.length === 0) throw new Error("No readable slides were found in this PPTX.");

  return {
    title: parseTitle(pkg) || "Imported PowerPoint presentation",
    slideSize,
    slides,
    sections,
    imageAssets,
    warnings: pkg.warnings,
    metadata: {
      creator: getText(getPptxXml(pkg, "docProps/core.xml", "docProps/core.xml")),
      slideCount: slides.length,
    },
    statistics: {
      hiddenSlides: slides.filter((slide) => slide.hidden).length,
      textBlocks: slides.reduce((sum, slide) => sum + slide.elements.filter((element) => element.kind === "text").length, 0),
      shapes: slides.reduce((sum, slide) => sum + slide.elements.filter((element) => element.kind === "shape").length, 0),
      images: slides.reduce((sum, slide) => sum + slide.elements.filter((element) => element.kind === "image").length, 0),
      tables: slides.reduce((sum, slide) => sum + slide.elements.filter((element) => element.kind === "table").length, 0),
      charts: slides.reduce((sum, slide) => sum + slide.elements.filter((element) => element.kind === "chart").length, 0),
      speakerNotes: slides.filter((slide) => slide.notes.trim()).length,
      hyperlinks: slides.reduce((sum, slide) => sum + slide.hyperlinks.length, 0),
      rasterisedElements: 0,
      unsupportedElements: 0,
    },
  };
}

function getRelationshipAttr(element: Element, localName: string) {
  return Array.from(element.attributes).find((attr) => attr.localName === localName && attr.namespaceURI?.includes("/relationships"))?.value;
}

function parseSlide(pkg: PptxPackage, slidePath: string, slideIndex: number, slideSize: PptxSlideSize, imageAssets: PptxImageAsset[], imageByHash: Map<string, PptxImageAsset>): PptxSlide {
  const xml = getPptxXml(pkg, slidePath, slidePath);
  if (!xml) throw new Error(`PPTX slide part is missing: ${slidePath}`);
  const relsPath = getSlideRelsPath(slidePath);
  const rels = parsePptxRelationships(pkg, relsPath, getBasePath(slidePath));
  const cSld = getElementsByLocalName(xml, "cSld")[0];
  const background = parseBackground(cSld);
  const elements: PptxSlideElement[] = [];
  for (const shape of getElementsByLocalName(xml, "sp")) {
    const text = parseTextBlock(shape, slidePath);
    if (text) elements.push(text);
    else {
      const parsedShape = parseShape(shape, slidePath);
      if (parsedShape) elements.push(parsedShape);
    }
  }
  for (const pic of getElementsByLocalName(xml, "pic")) {
    const image = parseImageElement(pkg, pic, slidePath, rels, imageAssets, imageByHash);
    if (image) elements.push(image);
  }
  for (const frame of getElementsByLocalName(xml, "graphicFrame")) {
    const table = parseTable(frame, slidePath);
    if (table) elements.push(table);
    else {
      const chart = parseChart(frame, slidePath);
      if (chart) {
        pkg.warnings.push({ code: "pptx-chart-fallback", category: "images", pageNumber: slideIndex, message: "A PowerPoint chart was imported as a labelled vector placeholder in this first phase." });
        elements.push(chart);
      }
    }
  }
  const title = detectSlideTitle(elements, slideIndex);
  const notes = parseNotes(pkg, rels);
  return {
    id: crypto.randomUUID(),
    sourceSlideId: getAttr(getElementsByLocalName(xml, "cSld")[0], "name") || `slide-${slideIndex}`,
    slideIndex,
    title,
    hidden: getAttr(getElementsByLocalName(xml, "sld")[0] ?? xml.documentElement, "show") === "0",
    sourcePath: slidePath,
    background,
    elements,
    notes,
    hyperlinks: parseHyperlinks(elements, rels, slideIndex),
  };
}

function parseSlideSize(xml: Document): PptxSlideSize {
  const sldSz = getElementsByLocalName(xml, "sldSz")[0];
  return {
    width: emuToPoints(getAttr(sldSz, "cx")) ?? 720,
    height: emuToPoints(getAttr(sldSz, "cy")) ?? 405,
  };
}

function parseTextBlock(shape: Element, sourcePath: string) {
  const txBody = getFirstChildByLocalName(shape, "txBody");
  if (!txBody) return null;
  const transform = parseTransform(shape);
  const paragraphs = getElementsByLocalName(txBody, "p").map((paragraph) => {
    const runs = getElementsByLocalName(paragraph, "r").map((run): PptxTextRun => ({
      text: getText(run),
      bold: getAttr(getFirstChildByLocalName(run, "rPr"), "b") === "1",
      italic: getAttr(getFirstChildByLocalName(run, "rPr"), "i") === "1",
      underline: Boolean(getAttr(getFirstChildByLocalName(run, "rPr"), "u")),
      strike: Boolean(getAttr(getFirstChildByLocalName(run, "rPr"), "strike")),
      fontSize: pptFontSize(getAttr(getFirstChildByLocalName(run, "rPr"), "sz")),
      color: parseColor(run),
    })).filter((run) => run.text);
    return runs.length ? runs : [{ text: getText(paragraph) }];
  }).filter((paragraph) => paragraph.some((run) => run.text.trim()));
  if (paragraphs.length === 0) return null;
  return { id: crypto.randomUUID(), kind: "text" as const, ...transform, paragraphs, alignment: "left" as const, sourcePath };
}

function parseShape(shape: Element, sourcePath: string) {
  const transform = parseTransform(shape);
  if (transform.width <= 0 || transform.height <= 0) return null;
  return {
    id: crypto.randomUUID(),
    kind: "shape" as const,
    preset: getAttr(getFirstChildByLocalName(getFirstChildByLocalName(shape, "prstGeom") ?? shape, "prstGeom"), "prst") ?? getAttr(getElementsByLocalName(shape, "prstGeom")[0], "prst") ?? "rect",
    ...transform,
    fill: parseColor(getFirstChildByLocalName(shape, "spPr") ?? shape) ?? "#f8fafc",
    stroke: "#94a3b8",
    strokeWidth: 1,
    sourcePath,
  };
}

function parseImageElement(pkg: PptxPackage, pic: Element, sourcePath: string, rels: Map<string, PptxRelationship>, imageAssets: PptxImageAsset[], imageByHash: Map<string, PptxImageAsset>) {
  const blip = getElementsByLocalName(pic, "blip")[0];
  const relationshipId = getAttr(blip, "embed") ?? getAttr(blip, "link");
  if (!relationshipId) return null;
  const rel = rels.get(relationshipId);
  if (!rel) return null;
  const bytes = pkg.files[rel.target];
  if (!bytes) {
    pkg.warnings.push({ code: "pptx-missing-image", category: "images", message: `Image relationship ${relationshipId} points to a missing file.` });
    return null;
  }
  const mimeType = getImageMimeType(rel.target);
  const hash = quickContentHash(bytes);
  let asset = imageByHash.get(hash);
  if (!asset) {
    const size = getImageSize(bytes, mimeType);
    asset = {
      id: crypto.randomUUID(),
      relationshipId,
      name: rel.target.split("/").pop() ?? "PowerPoint image",
      path: rel.target,
      mimeType,
      bytes,
      dataUrl: `data:${mimeType};base64,${bytesToBase64(bytes)}`,
      width: size.width,
      height: size.height,
      used: false,
    };
    imageByHash.set(hash, asset);
    imageAssets.push(asset);
  }
  return { id: crypto.randomUUID(), kind: "image" as const, relationshipId, assetId: asset.id, altText: getAttr(getElementsByLocalName(pic, "cNvPr")[0], "descr"), ...parseTransform(pic), sourcePath };
}

function parseTable(frame: Element, sourcePath: string) {
  const tbl = getElementsByLocalName(frame, "tbl")[0];
  if (!tbl) return null;
  const rows = getElementsByLocalName(tbl, "tr").map((row) => getElementsByLocalName(row, "tc").map((cell) => getText(cell).trim()));
  return { id: crypto.randomUUID(), kind: "table" as const, ...parseTransform(frame), rows, sourcePath };
}

function parseChart(frame: Element, sourcePath: string) {
  const chart = getElementsByLocalName(frame, "chart")[0];
  if (!chart) return null;
  return { id: crypto.randomUUID(), kind: "chart" as const, ...parseTransform(frame), chartType: "chart", sourcePath };
}

function parseTransform(element: Element) {
  const xfrm = getElementsByLocalName(element, "xfrm")[0];
  const off = getFirstChildByLocalName(xfrm ?? element, "off");
  const ext = getFirstChildByLocalName(xfrm ?? element, "ext");
  return {
    x: emuToPoints(getAttr(off, "x")) ?? 0,
    y: emuToPoints(getAttr(off, "y")) ?? 0,
    width: emuToPoints(getAttr(ext, "cx")) ?? 160,
    height: emuToPoints(getAttr(ext, "cy")) ?? 48,
    rotation: angleToDegrees(getAttr(xfrm, "rot")),
  };
}

function parseBackground(cSld?: Element) {
  const bg = cSld ? getElementsByLocalName(cSld, "bg")[0] : undefined;
  return bg ? parseColor(bg) : undefined;
}

function parseColor(element?: Element | null): string | undefined {
  if (!element) return undefined;
  const srgb = getElementsByLocalName(element, "srgbClr")[0];
  const value = getAttr(srgb, "val");
  return value && value.length === 6 ? `#${value}` : undefined;
}

function parseNotes(pkg: PptxPackage, rels: Map<string, PptxRelationship>) {
  const notesRel = Array.from(rels.values()).find((rel) => rel.type.includes("/notesSlide"));
  if (!notesRel) return "";
  return getText(getPptxXml(pkg, notesRel.target, notesRel.target)).trim();
}

function parseHyperlinks(elements: PptxSlideElement[], rels: Map<string, PptxRelationship>, slideIndex: number) {
  const links = [];
  for (const element of elements) {
    if (element.kind !== "text") continue;
    for (const paragraph of element.paragraphs) {
      const text = paragraph.map((run) => run.text).join(" ").trim();
      const rel = Array.from(rels.entries()).find(([, item]) => item.type.includes("/hyperlink"));
      if (rel && text) links.push({ slideIndex, elementId: element.id, text, url: rel[1].target, x: element.x, y: element.y, width: element.width, height: element.height });
    }
  }
  return links;
}

function parseSections(xml: Document) {
  return getElementsByLocalName(xml, "section").map((section) => ({
    id: getAttr(section, "id") ?? crypto.randomUUID(),
    name: getAttr(section, "name") ?? "Section",
    firstSlideId: getAttr(getElementsByLocalName(section, "sldId")[0], "id"),
  }));
}

function detectSlideTitle(elements: PptxSlideElement[], slideIndex: number) {
  const firstText = elements.find((element) => element.kind === "text");
  if (firstText?.kind === "text") return firstText.paragraphs.flat().map((run) => run.text).join(" ").trim().slice(0, 96) || `Slide ${slideIndex}`;
  return `Slide ${slideIndex}`;
}

function parseTitle(pkg: PptxPackage) {
  return getText(getPptxXml(pkg, "docProps/core.xml", "docProps/core.xml")).trim();
}

function getSlideRelsPath(slidePath: string) {
  const name = slidePath.split("/").pop() ?? "";
  return `${getBasePath(slidePath)}/_rels/${name}.rels`;
}

function getBasePath(path: string) {
  return path.split("/").slice(0, -1).join("/");
}

function emuToPoints(value?: string) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric / 12700 : undefined;
}

function angleToDegrees(value?: string) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric / 60000 : undefined;
}

function pptFontSize(value?: string) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric / 100 : undefined;
}

function getImageMimeType(path: string): PptxImageAsset["mimeType"] {
  const lower = path.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".gif")) return "image/gif";
  if (lower.endsWith(".svg")) return "image/svg+xml";
  return "image/unsupported";
}

function getImageSize(bytes: Uint8Array, mimeType: PptxImageAsset["mimeType"]) {
  if (mimeType === "image/png" && bytes.length > 24) return { width: readUint32(bytes, 16), height: readUint32(bytes, 20) };
  if (mimeType === "image/jpeg") return readJpegSize(bytes);
  if (mimeType === "image/svg+xml") {
    const text = strFromU8(bytes);
    return { width: Number(text.match(/\bwidth=["']?([\d.]+)/i)?.[1]) || 320, height: Number(text.match(/\bheight=["']?([\d.]+)/i)?.[1]) || 180 };
  }
  return { width: 320, height: 180 };
}

function readJpegSize(bytes: Uint8Array) {
  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) break;
    const marker = bytes[offset + 1];
    const length = (bytes[offset + 2] << 8) + bytes[offset + 3];
    if (marker >= 0xc0 && marker <= 0xc3) return { height: (bytes[offset + 5] << 8) + bytes[offset + 6], width: (bytes[offset + 7] << 8) + bytes[offset + 8] };
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
