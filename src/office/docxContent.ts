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
import type { DocxBlock, DocxImportOptions, DocxImportWarning, DocxInternalLink, DocxNote, DocxNoteReference, DocxPageSize, DocxParagraphBlock, DocxSectionNumbering, DocxTextRun, DocxTrackedChange, DocxWordBookmark, DocxWordComment } from "./types";
import { mergeParagraphFormat, mergeRunFormat, parseParagraphFormat, parseRunFormat, type DocxNumberingMap, type DocxStyleMap } from "./docxStyles";

type ParseState = {
  sourceOrder: number;
  listCounters: Map<string, number>;
  hyperlinks: Array<{ text: string; url: string }>;
  internalLinks: DocxInternalLink[];
  wordBookmarks: DocxWordBookmark[];
  wordComments: DocxWordComment[];
  trackedChanges: DocxTrackedChange[];
  trackedChangesDetected: number;
};

export function parseDocxContent(
  pkg: DocxPackage,
  styles: DocxStyleMap,
  numbering: DocxNumberingMap,
  relationships: DocxRelationships,
  warnings: DocxImportWarning[],
  options: Pick<DocxImportOptions, "fallbackPageSize" | "trackedChangesMode" | "fidelityMode">
) {
  const xml = getXml(pkg, "word/document.xml", "word/document.xml");
  if (!xml) throw new Error("The DOCX document XML is missing.");
  const body = getFirstChildByLocalName(xml.documentElement, "body");
  if (!body) throw new Error("The DOCX document body is missing.");
  const state: ParseState = {
    sourceOrder: 0,
    listCounters: new Map(),
    hyperlinks: [],
    internalLinks: [],
    wordBookmarks: [],
    wordComments: parseWordComments(pkg),
    trackedChanges: [],
    trackedChangesDetected: 0,
  };
  const blocks: DocxBlock[] = [];
  let paragraphCount = 0;
  let tableCount = 0;
  let listCount = 0;
  let sectionCount = 0;

  for (const child of Array.from(body.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const element = child as Element;
    if (element.localName === "p") {
      const parsed = parseParagraph(element, styles, numbering, relationships, state, warnings, options);
      if (parsed.pageBreakBefore) blocks.push(createPageBreak(state));
      if (parsed.paragraph.runs.length > 0 || parsed.paragraph.headingLevel || parsed.paragraph.list) {
        paragraphCount += 1;
        if (parsed.paragraph.list) listCount += 1;
        blocks.push(parsed.paragraph);
      }
      blocks.push(...parsed.images);
      for (let index = 0; index < parsed.pageBreaksAfter; index += 1) blocks.push(createPageBreak(state));
    }
    if (element.localName === "tbl") {
      tableCount += 1;
      blocks.push(parseTable(element, styles, numbering, relationships, state, warnings, options));
    }
    if (element.localName === "sectPr") sectionCount += 1;
  }

  sectionCount = Math.max(sectionCount, getElementsByLocalName(body, "sectPr").length || 1);
  const notes = [...parseNotes(pkg, "footnote"), ...parseNotes(pkg, "endnote")];
  const noteIds = new Set(notes.map((note) => `${note.kind}:${note.id}`));
  const noteReferences = blocks.flatMap((block) => (block.type === "paragraph" ? block.noteReferences ?? [] : []));
  const brokenNoteReferences = noteReferences.filter((reference) => !noteIds.has(`${reference.kind}:${reference.id}`)).length;
  if (brokenNoteReferences > 0) warnings.push({ code: "broken-note-reference", category: "footnotes", message: `${brokenNoteReferences} Word note references could not be resolved to footnote or endnote text.` });
  if (notes.length > 0) warnings.push({ code: "notes-imported", category: "footnotes", message: `${notes.filter((note) => note.kind === "footnote").length} footnotes and ${notes.filter((note) => note.kind === "endnote").length} endnotes were imported with bottom-of-page placement where space allows.` });
  if (state.wordComments.length > 0) warnings.push({ code: "word-comments-imported", category: "comments", message: `${state.wordComments.length} Word comments were imported into Publish Pro review comments with approximate source mapping where exact ranges are unavailable.` });
  if (state.trackedChangesDetected > 0) warnings.push({ code: "tracked-changes-mode", category: "trackedChanges", message: `${state.trackedChangesDetected} tracked-change elements were detected. Import used ${getTrackedChangesModeLabel(options.trackedChangesMode)} mode.` });
  if (state.internalLinks.length > 0) warnings.push({ code: "internal-links-detected", category: "hyperlinks", message: `${state.internalLinks.length} internal Word hyperlinks were detected and will be mapped to PDF page destinations where possible.` });

  const sectionNumbering = parseSectionNumbering(body);
  if (sectionNumbering.length > 0) warnings.push({ code: "section-numbering-imported", category: "sections", message: `${sectionNumbering.length} Word section page-numbering starts were mapped into Publish Pro numbering sections where page mapping is reliable.` });

  return {
    blocks,
    pageSize: parsePageSize(body) ?? getFallbackPageSize(options.fallbackPageSize),
    hyperlinks: state.hyperlinks,
    notes,
    wordBookmarks: state.wordBookmarks,
    internalLinks: state.internalLinks,
    wordComments: state.wordComments,
    trackedChanges: state.trackedChanges,
    sectionNumbering,
    statistics: {
      paragraphCount,
      tableCount,
      listCount,
      imageCount: 0,
      footnoteCount: notes.filter((note) => note.kind === "footnote").length,
      endnoteCount: notes.filter((note) => note.kind === "endnote").length,
      commentsDetected: state.wordComments.length,
      trackedChangesDetected: state.trackedChangesDetected,
      sectionCount,
      notesPlacedExactly: 0,
      notesUsingFallback: notes.length,
      brokenNoteReferences,
      internalLinksDetected: state.internalLinks.length,
      internalLinksMapped: 0,
      commentsImported: state.wordComments.length,
      approximateComments: state.wordComments.filter((comment) => comment.approximate).length,
    },
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
  state: ParseState,
  warnings: DocxImportWarning[],
  options: Pick<DocxImportOptions, "trackedChangesMode" | "fidelityMode">
) {
  const sourceOrder = nextOrder(state);
  const pPr = getFirstChildByLocalName(paragraph, "pPr");
  const styleId = getAttr(getFirstChildByLocalName(pPr ?? paragraph, "pStyle"), "val") ?? styles.defaultParagraphStyleId;
  const style = styles.resolveParagraphStyle(styleId);
  const directParagraph = parseParagraphFormat(pPr);
  const numberingInfo = parseList(pPr, numbering, state.listCounters);
  const numberingParagraph = numberingInfo ? numbering.get(numberingInfo.numId)?.get(String(numberingInfo.level))?.paragraph : undefined;
  const paragraphFormat = mergeParagraphFormat(styles.documentDefaults.paragraph, style?.paragraph, numberingParagraph, directParagraph);
  const paragraphBlock: DocxParagraphBlock = {
    id: `docx-block-${sourceOrder}`,
    type: "paragraph",
    sourcePath: "word/document.xml",
    sourceOrder,
    runs: [],
    styleId,
    styleName: style?.name,
    headingLevel: style?.headingLevel,
    alignment: paragraphFormat.alignment,
    indentLeft: paragraphFormat.indentLeft,
    indentRight: paragraphFormat.indentRight,
    firstLineIndent: paragraphFormat.firstLineIndent,
    hangingIndent: paragraphFormat.hangingIndent,
    spacingBefore: paragraphFormat.spacingBefore,
    spacingAfter: paragraphFormat.spacingAfter,
    lineSpacing: paragraphFormat.lineSpacing,
    keepWithNext: paragraphFormat.keepWithNext,
    keepLinesTogether: paragraphFormat.keepLinesTogether,
    pageBreakBefore: paragraphFormat.pageBreakBefore,
    widowControl: paragraphFormat.widowControl,
    list: numberingInfo,
  };
  const sourceBookmarkNames = getElementsByLocalName(paragraph, "bookmarkStart").map((bookmark) => getAttr(bookmark, "name")).filter((name): name is string => typeof name === "string" && !name.startsWith("_"));
  const sourceCommentIds = [
    ...getElementsByLocalName(paragraph, "commentRangeStart").map((comment) => getAttr(comment, "id")),
    ...getElementsByLocalName(paragraph, "commentReference").map((comment) => getAttr(comment, "id")),
  ].filter((id): id is string => Boolean(id));
  if (sourceBookmarkNames.length > 0) {
    paragraphBlock.wordBookmarkNames = Array.from(new Set(sourceBookmarkNames));
    paragraphBlock.wordBookmarkNames.forEach((name) => state.wordBookmarks.push({ name, blockId: paragraphBlock.id, sourceOrder }));
  }
  if (sourceCommentIds.length > 0) {
    paragraphBlock.wordCommentIds = Array.from(new Set(sourceCommentIds));
    const sourceText = getText(paragraph).trim();
    for (const commentId of paragraphBlock.wordCommentIds) {
      const existing = state.wordComments.find((comment) => comment.sourceId === commentId);
      if (existing) {
        existing.blockId = paragraphBlock.id;
        existing.sourceText = sourceText || existing.sourceText;
        existing.approximate = false;
      }
    }
  }
  const images: DocxBlock[] = [];
  let pageBreaksAfter = 0;

  for (const child of Array.from(paragraph.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const element = child as Element;
    if (element.localName === "r") {
      const parsed = parseRun(element, styles, relationships, state, undefined, warnings);
      paragraphBlock.runs.push(...parsed.runs);
      appendNoteReferences(paragraphBlock, parsed.noteReferences);
      images.push(...parsed.images);
      pageBreaksAfter += parsed.pageBreaks;
    }
    if (element.localName === "hyperlink") {
      const relationshipId = getAttr(element, "id");
      const anchor = getAttr(element, "anchor");
      const target = getHyperlinkTarget(relationships, relationshipId) ?? anchor;
      const runs = getChildrenByLocalName(element, "r").flatMap((run) => parseRun(run, styles, relationships, state, target, warnings).runs);
      paragraphBlock.runs.push(...runs);
      const text = runs.map((run) => run.text).join("").trim();
      if (target && text) state.hyperlinks.push({ text, url: target });
      if (anchor && text) state.internalLinks.push({ blockId: paragraphBlock.id, anchor, text, sourceOrder });
    }
    if (element.localName === "ins" || element.localName === "del" || element.localName === "moveFrom" || element.localName === "moveTo") {
      state.trackedChangesDetected += 1;
      const changeType = getTrackedChangeType(element.localName);
      const changeText = getText(element).trim();
      state.trackedChanges.push({
        id: getAttr(element, "id") ?? `tracked-${state.trackedChangesDetected}`,
        type: changeType,
        author: getAttr(element, "author"),
        date: getAttr(element, "date"),
        text: changeText,
        blockId: paragraphBlock.id,
      });
      if (!shouldRenderTrackedChange(element.localName, options.trackedChangesMode)) continue;
      for (const run of getChildrenByLocalName(element, "r")) {
        const parsed = parseRun(run, styles, relationships, state, undefined, warnings);
        paragraphBlock.runs.push(...parsed.runs);
        appendNoteReferences(paragraphBlock, parsed.noteReferences);
        pageBreaksAfter += parsed.pageBreaks;
      }
    }
  }
  return { paragraph: paragraphBlock, images, pageBreakBefore: Boolean(paragraphFormat.pageBreakBefore), pageBreaksAfter };
}

function parseRun(run: Element, styles: DocxStyleMap, relationships: DocxRelationships, state: ParseState, hyperlink: string | undefined, warnings: DocxImportWarning[]): { runs: DocxTextRun[]; images: DocxBlock[]; pageBreaks: number; noteReferences: DocxNoteReference[] } {
  const rPr = getFirstChildByLocalName(run, "rPr");
  const characterStyleId = getAttr(getFirstChildByLocalName(rPr ?? run, "rStyle"), "val");
  const characterStyle = styles.resolveCharacterStyle(characterStyleId);
  const base = { ...mergeRunFormat(styles.documentDefaults.run, characterStyle?.run, parseRunFormat(rPr)), hyperlink };
  const runs: DocxTextRun[] = [];
  const images: DocxBlock[] = [];
  const noteReferences: DocxNoteReference[] = [];
  let pageBreaks = 0;
  for (const child of Array.from(run.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const element = child as Element;
    if (element.localName === "t") runs.push(createRun(state, base, element.textContent ?? ""));
    if (element.localName === "tab") runs.push(createRun(state, base, "\t"));
    if (element.localName === "br") {
      if (getAttr(element, "type") === "page") pageBreaks += 1;
      else runs.push(createRun(state, base, "\n"));
    }
    if (element.localName === "drawing" || element.localName === "pict") {
      const relationshipId = getElementsByLocalName(element, "blip").map((blip) => getAttr(blip, "embed") ?? getAttr(blip, "link")).find(Boolean);
      const extent = getElementsByLocalName(element, "extent")[0];
      const docPr = getElementsByLocalName(element, "docPr")[0];
      if (relationshipId) {
        images.push({
          id: `docx-block-${nextOrder(state)}`,
          type: "image",
          sourcePath: "word/document.xml",
          sourceOrder: state.sourceOrder,
          relationshipId,
          altText: getAttr(docPr, "descr") ?? getAttr(docPr, "title"),
          width: emuToPoints(getAttr(extent, "cx")),
          height: emuToPoints(getAttr(extent, "cy")),
          floating: getElementsByLocalName(element, "anchor").length > 0,
          wrapMode: getElementsByLocalName(element, "wrapSquare").length > 0 ? "square" : getElementsByLocalName(element, "wrapTopAndBottom").length > 0 ? "topBottom" : "inline",
        });
      } else {
        warnings.push({ code: "unsupported-image-reference", message: "An image or drawing was found without a resolvable relationship ID and was skipped." });
      }
    }
    if (element.localName === "footnoteReference" || element.localName === "endnoteReference") {
      const kind = element.localName === "endnoteReference" ? "endnote" : "footnote";
      const id = getAttr(element, "id") ?? "";
      const marker = id ? `[${id}]` : "[]";
      noteReferences.push({ id, kind, marker });
      runs.push(createRun(state, { ...base, verticalAlign: "superscript" }, marker));
    }
    if (element.localName === "fldChar" || element.localName === "instrText") {
      const field = (element.textContent ?? "").trim();
      if (field) warnings.push({ code: "word-field", message: `Word field detected: ${field}. Visible field result text is imported where available.` });
    }
  }
  return { runs, images, pageBreaks, noteReferences };
}

function parseTable(table: Element, styles: DocxStyleMap, numbering: DocxNumberingMap, relationships: DocxRelationships, state: ParseState, warnings: DocxImportWarning[], options: Pick<DocxImportOptions, "trackedChangesMode" | "fidelityMode">) {
  const sourceOrder = nextOrder(state);
  const grid = getFirstChildByLocalName(table, "tblGrid");
  const columnWidths = grid ? getChildrenByLocalName(grid, "gridCol").map((column) => twipsToPoints(getAttr(column, "w")) ?? 0).filter((width) => width > 0) : undefined;
  const preferredWidth = twipsToPoints(getAttr(getFirstChildByLocalName(getFirstChildByLocalName(table, "tblPr") ?? table, "tblW"), "w"));
  const rows = getTableRows(table).map((row, rowIndex) =>
    getTableCells(row).map((cell) => {
      const tcPr = getFirstChildByLocalName(cell, "tcPr");
      const gridSpan = Number(getAttr(getFirstChildByLocalName(tcPr ?? cell, "gridSpan"), "val") ?? 1) || 1;
      if (gridSpan > 1) warnings.push({ code: "merged-cell-simplified", message: "A merged table cell was imported using simplified column geometry." });
      return {
        blocks: getChildrenByLocalName(cell, "p").map((paragraph) => parseParagraph(paragraph, styles, numbering, relationships, state, warnings, options).paragraph),
        background: normalizeColor(getAttr(getFirstChildByLocalName(tcPr ?? cell, "shd"), "fill")),
        gridSpan,
        width: twipsToPoints(getAttr(getFirstChildByLocalName(tcPr ?? cell, "tcW"), "w")),
        verticalAlign: parseVerticalAlign(getAttr(getFirstChildByLocalName(tcPr ?? cell, "vAlign"), "val")),
      };
    }).map((cell, cellIndex) => ({ ...cell, headerRow: rowIndex === 0 && cellIndex === 0 && Boolean(getFirstChildByLocalName(getFirstChildByLocalName(row, "trPr") ?? row, "tblHeader")) }))
  );
  return {
    id: `docx-block-${sourceOrder}`,
    type: "table" as const,
    sourcePath: "word/document.xml",
    sourceOrder,
    rows,
    columnWidths,
    preferredWidth,
    headerRows: getTableRows(table).map((row, index) => (getFirstChildByLocalName(getFirstChildByLocalName(row, "trPr") ?? row, "tblHeader") ? index : -1)).filter((index) => index >= 0),
  };
}

function parseList(pPr: Element | undefined, numbering: DocxNumberingMap, listCounters: Map<string, number>): DocxParagraphBlock["list"] {
  if (!pPr) return undefined;
  const numPr = getFirstChildByLocalName(pPr, "numPr");
  if (!numPr) return undefined;
  const numId = getAttr(getFirstChildByLocalName(numPr, "numId"), "val");
  const ilvl = getAttr(getFirstChildByLocalName(numPr, "ilvl"), "val") ?? "0";
  if (!numId) return undefined;
  const level = Number(ilvl) || 0;
  const definition = numbering.get(numId)?.get(ilvl);
  const key = `${numId}:${ilvl}`;
  const index = (listCounters.get(key) ?? ((definition?.start ?? 1) - 1)) + 1;
  listCounters.set(key, index);
  return {
    numId,
    level,
    kind: definition?.kind ?? "number",
    format: definition?.format,
    text: definition?.text,
    index,
  };
}

function createRun(state: ParseState, base: Omit<DocxTextRun, "id" | "text">, text: string): DocxTextRun {
  const nextText = base.allCaps ? text.toUpperCase() : text.replace(/\u00a0/g, " ");
  return { id: `docx-run-${nextOrder(state)}`, ...base, text: nextText };
}

function createPageBreak(state: ParseState): DocxBlock {
  const sourceOrder = nextOrder(state);
  return { id: `docx-block-${sourceOrder}`, type: "pageBreak", sourcePath: "word/document.xml", sourceOrder };
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

function normalizeColor(value?: string) {
  if (!value || value === "auto" || value.length !== 6) return undefined;
  return `#${value}`;
}

function twipsToPoints(value?: string) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric / 20 : undefined;
}

function emuToPoints(value?: string) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric / 12700 : undefined;
}

function parseVerticalAlign(value?: string): "top" | "middle" | "bottom" {
  if (value === "center") return "middle";
  if (value === "bottom") return "bottom";
  return "top";
}

function uniqueJoin(values: string[]) {
  return Array.from(new Set(values.filter(Boolean))).join(" · ");
}

function nextOrder(state: ParseState) {
  state.sourceOrder += 1;
  return state.sourceOrder;
}

function parseNotes(pkg: DocxPackage, kind: "footnote" | "endnote"): DocxNote[] {
  const xml = getXml(pkg, kind === "footnote" ? "word/footnotes.xml" : "word/endnotes.xml", kind === "footnote" ? "word/footnotes.xml" : "word/endnotes.xml");
  if (!xml) return [];
  const relationships = parsePartRelationships(pkg, kind === "footnote" ? "word/_rels/footnotes.xml.rels" : "word/_rels/endnotes.xml.rels");
  return getElementsByLocalName(xml, kind)
    .filter((note) => Number(getAttr(note, "id")) > 0)
    .map((note) => {
      const id = getAttr(note, "id") ?? "";
      return {
        id,
        kind,
        marker: `[${id}]`,
        text: getText(note).trim(),
        hyperlinks: getElementsByLocalName(note, "hyperlink").map((link) => {
          const relationshipId = getAttr(link, "id");
          return { text: getText(link).trim(), url: relationships.get(relationshipId ?? "") ?? getAttr(link, "anchor") ?? "" };
        }).filter((link) => link.text && link.url),
      };
    });
}

function parsePartRelationships(pkg: DocxPackage, path: string) {
  const xml = getXml(pkg, path, path);
  if (!xml) return new Map<string, string>();
  const entries: Array<[string, string]> = getElementsByLocalName(xml, "Relationship").flatMap((relationship) => {
    const id = getAttr(relationship, "Id");
    const target = getAttr(relationship, "Target");
    return id && target ? [[id, target]] : [];
  });
  return new Map(entries);
}

function parseWordComments(pkg: DocxPackage): DocxWordComment[] {
  const xml = getXml(pkg, "word/comments.xml", "word/comments.xml");
  if (!xml) return [];
  return getElementsByLocalName(xml, "comment").map((comment) => ({
    sourceId: getAttr(comment, "id") ?? crypto.randomUUID(),
    author: getAttr(comment, "author") || "Word reviewer",
    initials: getAttr(comment, "initials"),
    date: getAttr(comment, "date"),
    text: getText(comment).trim(),
    approximate: true,
  })).filter((comment) => comment.text);
}

function parseSectionNumbering(body: Element): DocxSectionNumbering[] {
  return getElementsByLocalName(body, "sectPr").flatMap((section) => {
    const pgNumType = getFirstChildByLocalName(section, "pgNumType");
    if (!pgNumType) return [];
    const startValue = Number(getAttr(pgNumType, "start") ?? 1) || 1;
    return [{
      blockId: `docx-section-${getElementsByLocalName(body, "sectPr").indexOf(section) + 1}`,
      sourceOrder: getElementsByLocalName(body, "sectPr").indexOf(section) + 1,
      startValue,
      format: parsePageNumberFormat(getAttr(pgNumType, "fmt")),
      prefix: getAttr(pgNumType, "chapSep") ? `${getAttr(pgNumType, "chapSep")}` : "",
      unnumbered: getAttr(pgNumType, "start") === "0",
    }];
  });
}

function parsePageNumberFormat(value?: string): DocxSectionNumbering["format"] {
  if (value === "lowerRoman") return "romanLower";
  if (value === "upperRoman") return "romanUpper";
  if (value === "lowerLetter") return "alphaLower";
  if (value === "upperLetter") return "alphaUpper";
  return "decimal";
}

function appendNoteReferences(paragraph: DocxParagraphBlock, references: DocxNoteReference[]) {
  if (references.length === 0) return;
  paragraph.noteReferences = [...(paragraph.noteReferences ?? []), ...references];
}

function shouldRenderTrackedChange(localName: string, mode: DocxImportOptions["trackedChangesMode"]) {
  const normalizedMode = normalizeTrackedChangesMode(mode);
  if (localName === "del" || localName === "moveFrom") return normalizedMode === "rejectAll";
  if (localName === "ins" || localName === "moveTo") return normalizedMode !== "rejectAll";
  return true;
}

function normalizeTrackedChangesMode(mode: DocxImportOptions["trackedChangesMode"] | "accepted" | "rejectDeletions") {
  if (mode === "accepted") return "acceptAll";
  if (mode === "rejectDeletions") return "rejectAll";
  return mode;
}

function getTrackedChangesModeLabel(mode: DocxImportOptions["trackedChangesMode"]) {
  const normalizedMode = normalizeTrackedChangesMode(mode);
  if (normalizedMode === "rejectAll") return "Reject all";
  if (normalizedMode === "summary") return "Preserve review summary";
  return "Accept all";
}

function getTrackedChangeType(localName: string): DocxTrackedChange["type"] {
  if (localName === "del") return "deletion";
  if (localName === "moveFrom") return "moveFrom";
  if (localName === "moveTo") return "moveTo";
  return "insertion";
}
