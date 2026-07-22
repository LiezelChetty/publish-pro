import {
  getAttr,
  getChildrenByLocalName,
  getElementsByLocalName,
  getFirstChildByLocalName,
  getPackageText,
  getText,
  getXml,
  parseXml,
  type DocxPackage,
} from "./docxPackage";
import { getHyperlinkTarget, type DocxRelationships } from "./docxImages";
import { getTableCells, getTableRows } from "./docxTables";
import type { DocxBlock, DocxImportWarning, DocxPageSize, DocxParagraphBlock, DocxTextRun } from "./types";
import type { DocxNumberingMap, DocxStyleMap } from "./docxStyles";

export function parseDocxContent(pkg: DocxPackage, styles: DocxStyleMap, numbering: DocxNumberingMap, relationships: DocxRelationships, warnings: DocxImportWarning[], fallbackPageSize: "a4" | "letter" | "legal" = "letter") {
  const xml = getXml(pkg, "word/document.xml", "word/document.xml");
  if (!xml) throw new Error("The DOCX document XML is missing.");
  const body = getFirstChildByLocalName(xml.documentElement, "body");
  if (!body) throw new Error("The DOCX document body is missing.");
  const listCounters = new Map<string, number>();
  const blocks: DocxBlock[] = [];
  let paragraphCount = 0;
  let tableCount = 0;
  let listCount = 0;
  const hyperlinks: Array<{ text: string; url: string }> = [];

  for (const child of Array.from(body.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const element = child as Element;
    if (element.localName === "p") {
      const parsed = parseParagraph(element, styles, numbering, relationships, listCounters, hyperlinks);
      if (parsed.pageBreakBefore) blocks.push({ type: "pageBreak" });
      if (parsed.paragraph.runs.length > 0 || parsed.paragraph.headingLevel || parsed.paragraph.list) {
        paragraphCount += 1;
        if (parsed.paragraph.list) listCount += 1;
        blocks.push(parsed.paragraph);
      }
      blocks.push(...parsed.images);
      for (let index = 0; index < parsed.pageBreaksAfter; index += 1) blocks.push({ type: "pageBreak" });
    }
    if (element.localName === "tbl") {
      tableCount += 1;
      blocks.push(parseTable(element, styles, numbering, relationships, listCounters, hyperlinks));
    }
  }

  const footnoteCount = countFootnotes(pkg);
  if (footnoteCount > 0) warnings.push({ code: "footnotes-fallback", message: `${footnoteCount} footnote references were detected. Footnote text is imported as readable document text where present, but exact Word footnote pagination is not yet guaranteed.` });
  return {
    blocks,
    pageSize: parsePageSize(body) ?? getFallbackPageSize(fallbackPageSize),
    hyperlinks,
    statistics: { paragraphCount, tableCount, listCount, imageCount: 0, footnoteCount },
  };
}

export function parseHeaderFooterText(pkg: DocxPackage, relationships: DocxRelationships) {
  const headers: string[] = [];
  const footers: string[] = [];
  for (const relationship of relationships.targets.values()) {
    if (relationship.type.includes("/header")) {
      const text = getText(parseOptionalXml(pkg, relationship.target));
      if (text.trim()) headers.push(text.trim());
    }
    if (relationship.type.includes("/footer")) {
      const text = getText(parseOptionalXml(pkg, relationship.target));
      if (text.trim()) footers.push(text.trim());
    }
  }
  return { headerText: uniqueJoin(headers), footerText: uniqueJoin(footers) };
}

function parseOptionalXml(pkg: DocxPackage, path: string) {
  const text = getPackageText(pkg, path);
  return text ? parseXml(text, path) : null;
}

function parseParagraph(
  paragraph: Element,
  styles: DocxStyleMap,
  numbering: DocxNumberingMap,
  relationships: DocxRelationships,
  listCounters: Map<string, number>,
  hyperlinks: Array<{ text: string; url: string }>
) {
  const pPr = getFirstChildByLocalName(paragraph, "pPr");
  const styleId = getAttr(getFirstChildByLocalName(pPr ?? paragraph, "pStyle"), "val");
  const headingLevel = styleId ? styles.headingLevels.get(styleId) : undefined;
  const alignment = parseAlignment(getAttr(getFirstChildByLocalName(pPr ?? paragraph, "jc"), "val"));
  const indentLeft = twipsToPoints(getAttr(getFirstChildByLocalName(pPr ?? paragraph, "ind"), "left"));
  const spacing = getFirstChildByLocalName(pPr ?? paragraph, "spacing");
  const paragraphBlock: DocxParagraphBlock = {
    type: "paragraph",
    runs: [],
    styleId,
    headingLevel,
    alignment,
    indentLeft,
    spacingBefore: twipsToPoints(getAttr(spacing, "before")),
    spacingAfter: twipsToPoints(getAttr(spacing, "after")),
    list: parseList(pPr, numbering, listCounters),
  };
  const images: DocxBlock[] = [];
  let pageBreakBefore = false;
  let pageBreaksAfter = 0;
  const pageBreakBeforeValue = getAttr(getFirstChildByLocalName(pPr ?? paragraph, "pageBreakBefore"), "val");
  if (pageBreakBeforeValue !== undefined && pageBreakBeforeValue !== "0" && pageBreakBeforeValue !== "false") pageBreakBefore = true;

  for (const child of Array.from(paragraph.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const element = child as Element;
    if (element.localName === "r") {
      const parsed = parseRun(element);
      paragraphBlock.runs.push(...parsed.runs);
      images.push(...parsed.images);
      pageBreaksAfter += parsed.pageBreaks;
    }
    if (element.localName === "hyperlink") {
      const relationshipId = getAttr(element, "id");
      const target = getHyperlinkTarget(relationships, relationshipId);
      const runs = getChildrenByLocalName(element, "r").flatMap((run) => parseRun(run, target).runs);
      paragraphBlock.runs.push(...runs);
      const text = runs.map((run) => run.text).join("").trim();
      if (target && text) hyperlinks.push({ text, url: target });
    }
  }
  return { paragraph: paragraphBlock, images, pageBreakBefore, pageBreaksAfter };
}

function parseRun(run: Element, hyperlink?: string): { runs: DocxTextRun[]; images: DocxBlock[]; pageBreaks: number } {
  const rPr = getFirstChildByLocalName(run, "rPr");
  const base: Omit<DocxTextRun, "text"> = {
    bold: Boolean(getFirstChildByLocalName(rPr ?? run, "b")),
    italic: Boolean(getFirstChildByLocalName(rPr ?? run, "i")),
    underline: Boolean(getFirstChildByLocalName(rPr ?? run, "u")),
    strike: Boolean(getFirstChildByLocalName(rPr ?? run, "strike") || getFirstChildByLocalName(rPr ?? run, "dstrike")),
    color: normalizeColor(getAttr(getFirstChildByLocalName(rPr ?? run, "color"), "val")),
    fontSize: halfPointsToPoints(getAttr(getFirstChildByLocalName(rPr ?? run, "sz"), "val")),
    hyperlink,
  };
  const runs: DocxTextRun[] = [];
  const images: DocxBlock[] = [];
  let pageBreaks = 0;
  for (const child of Array.from(run.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const element = child as Element;
    if (element.localName === "t") runs.push({ ...base, text: element.textContent ?? "" });
    if (element.localName === "tab") runs.push({ ...base, text: "    " });
    if (element.localName === "br") {
      if (getAttr(element, "type") === "page") pageBreaks += 1;
      else runs.push({ ...base, text: "\n" });
    }
    if (element.localName === "drawing" || element.localName === "pict") {
      const relationshipId = getElementsByLocalName(element, "blip").map((blip) => getAttr(blip, "embed") ?? getAttr(blip, "link")).find(Boolean);
      if (relationshipId) images.push({ type: "image", relationshipId });
    }
    if (element.localName === "footnoteReference") runs.push({ ...base, text: `[${getAttr(element, "id") ?? ""}]` });
  }
  return { runs, images, pageBreaks };
}

function parseTable(table: Element, styles: DocxStyleMap, numbering: DocxNumberingMap, relationships: DocxRelationships, listCounters: Map<string, number>, hyperlinks: Array<{ text: string; url: string }>) {
  return {
    type: "table" as const,
    rows: getTableRows(table).map((row) =>
      getTableCells(row).map((cell) => ({
        blocks: getChildrenByLocalName(cell, "p").map((paragraph) => parseParagraph(paragraph, styles, numbering, relationships, listCounters, hyperlinks).paragraph),
        background: normalizeColor(getAttr(getFirstChildByLocalName(getFirstChildByLocalName(cell, "tcPr") ?? cell, "shd"), "fill")),
      }))
    ),
  };
}

function parseList(pPr: Element | undefined, numbering: DocxNumberingMap, listCounters: Map<string, number>) {
  if (!pPr) return undefined;
  const numPr = getFirstChildByLocalName(pPr, "numPr");
  if (!numPr) return undefined;
  const numId = getAttr(getFirstChildByLocalName(numPr, "numId"), "val");
  const ilvl = getAttr(getFirstChildByLocalName(numPr, "ilvl"), "val") ?? "0";
  if (!numId) return undefined;
  const key = `${numId}:${ilvl}`;
  const index = (listCounters.get(key) ?? 0) + 1;
  listCounters.set(key, index);
  return {
    level: Number(ilvl) || 0,
    kind: numbering.get(numId)?.get(ilvl) ?? "number",
    index,
  };
}

function parsePageSize(body: Element): DocxPageSize | undefined {
  const sectPr = getChildrenByLocalName(body, "sectPr")[0] ?? getElementsByLocalName(body, "sectPr")[0];
  const pgSz = getFirstChildByLocalName(sectPr ?? body, "pgSz");
  const pgMar = getFirstChildByLocalName(sectPr ?? body, "pgMar");
  if (!pgSz) return undefined;
  const width = twipsToPoints(getAttr(pgSz, "w")) ?? 612;
  const height = twipsToPoints(getAttr(pgSz, "h")) ?? 792;
  const orientation = getAttr(pgSz, "orient");
  return {
    width: orientation === "landscape" ? Math.max(width, height) : width,
    height: orientation === "landscape" ? Math.min(width, height) : height,
    marginTop: twipsToPoints(getAttr(pgMar, "top")) ?? 72,
    marginRight: twipsToPoints(getAttr(pgMar, "right")) ?? 72,
    marginBottom: twipsToPoints(getAttr(pgMar, "bottom")) ?? 72,
    marginLeft: twipsToPoints(getAttr(pgMar, "left")) ?? 72,
  };
}

export function getFallbackPageSize(size: "a4" | "letter" | "legal"): DocxPageSize {
  if (size === "a4") return { width: 595.28, height: 841.89, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72 };
  if (size === "legal") return { width: 612, height: 1008, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72 };
  return { width: 612, height: 792, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72 };
}

function countFootnotes(pkg: DocxPackage) {
  const xml = getXml(pkg, "word/footnotes.xml", "word/footnotes.xml");
  return xml ? getElementsByLocalName(xml, "footnote").filter((footnote) => Number(getAttr(footnote, "id")) > 0).length : 0;
}

function parseAlignment(value?: string): DocxParagraphBlock["alignment"] {
  if (value === "center") return "center";
  if (value === "right" || value === "end") return "right";
  return "left";
}

function normalizeColor(value?: string) {
  if (!value || value === "auto" || value.length !== 6) return undefined;
  return `#${value}`;
}

function twipsToPoints(value?: string) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric / 20 : undefined;
}

function halfPointsToPoints(value?: string) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric / 2 : undefined;
}

function uniqueJoin(values: string[]) {
  return Array.from(new Set(values.filter(Boolean))).join(" · ");
}
