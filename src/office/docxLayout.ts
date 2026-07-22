import { PDFArray, PDFHexString, PDFName, PDFNumber, PDFDocument, StandardFonts, rgb, type PDFFont, type PDFPage } from "pdf-lib";
import type { DocxBlock, DocxExtractedImage, DocxHeading, DocxImageBlock, DocxImportOptions, DocxIntermediateDocument, DocxParagraphBlock, DocxRenderedLink, DocxSourceMapping, DocxTableBlock, DocxTextRun } from "./types";

type LayoutContext = {
  pdf: PDFDocument;
  page: PDFPage;
  pageNumber: number;
  y: number;
  fontRegular: PDFFont;
  fontBold: PDFFont;
  fontItalic: PDFFont;
  fontBoldItalic: PDFFont;
  images: Map<string, DocxExtractedImage>;
  headings: DocxHeading[];
  links: DocxRenderedLink[];
  sourceMappings: DocxSourceMapping[];
  warnings: DocxIntermediateDocument["warnings"];
  fidelityMode: DocxImportOptions["fidelityMode"];
};

export async function renderDocxToPdf(document: DocxIntermediateDocument, fallbackFont: DocxImportOptions["fallbackFont"] = "Helvetica", fidelityMode: DocxImportOptions["fidelityMode"] = "balanced") {
  const pdf = await PDFDocument.create();
  const fontSet = getFontSet(fallbackFont);
  const fontRegular = await pdf.embedFont(fontSet.regular);
  const fontBold = await pdf.embedFont(fontSet.bold);
  const fontItalic = await pdf.embedFont(fontSet.italic);
  const fontBoldItalic = await pdf.embedFont(fontSet.boldItalic);
  const images = new Map(document.images.map((image) => [image.relationshipId, image]));
  const ctx: LayoutContext = {
    pdf,
    page: pdf.addPage([document.pageSize.width, document.pageSize.height]),
    pageNumber: 1,
    y: document.pageSize.height - document.pageSize.marginTop,
    fontRegular,
    fontBold,
    fontItalic,
    fontBoldItalic,
    images,
    headings: [],
    links: [],
    sourceMappings: [],
    warnings: document.warnings,
    fidelityMode,
  };

  for (const block of document.blocks) {
    if (block.type === "pageBreak") {
      addPage(ctx, document);
      continue;
    }
    if (block.type === "paragraph") drawParagraph(ctx, document, block);
    if (block.type === "image") await drawImage(ctx, document, block);
    if (block.type === "table") drawTable(ctx, document, block);
  }

  if (ctx.links.length > 0) addExternalLinkAnnotations(pdf, ctx.links);
  return { bytes: await pdf.save(), pageCount: pdf.getPageCount(), headings: ctx.headings, links: ctx.links, sourceMappings: ctx.sourceMappings };
}

function getFontSet(fallbackFont: DocxImportOptions["fallbackFont"]) {
  if (fallbackFont === "Times Roman") {
    return {
      regular: StandardFonts.TimesRoman,
      bold: StandardFonts.TimesRomanBold,
      italic: StandardFonts.TimesRomanItalic,
      boldItalic: StandardFonts.TimesRomanBoldItalic,
    };
  }
  if (fallbackFont === "Courier") {
    return {
      regular: StandardFonts.Courier,
      bold: StandardFonts.CourierBold,
      italic: StandardFonts.CourierOblique,
      boldItalic: StandardFonts.CourierBoldOblique,
    };
  }
  return {
    regular: StandardFonts.Helvetica,
    bold: StandardFonts.HelveticaBold,
    italic: StandardFonts.HelveticaOblique,
    boldItalic: StandardFonts.HelveticaBoldOblique,
  };
}

function addPage(ctx: LayoutContext, document: DocxIntermediateDocument) {
  ctx.page = ctx.pdf.addPage([document.pageSize.width, document.pageSize.height]);
  ctx.pageNumber += 1;
  ctx.y = document.pageSize.height - document.pageSize.marginTop;
}

function ensureSpace(ctx: LayoutContext, document: DocxIntermediateDocument, height: number, keepWithNext = false) {
  const minimumRemaining = keepWithNext ? Math.max(height, 72) : height;
  if (ctx.y - minimumRemaining < document.pageSize.marginBottom) addPage(ctx, document);
}

function drawParagraph(ctx: LayoutContext, document: DocxIntermediateDocument, paragraph: DocxParagraphBlock) {
  const text = getParagraphText(paragraph);
  if (!text.trim() && !paragraph.headingLevel) return;
  const firstRun = paragraph.runs[0];
  const fontSize = getParagraphFontSize(paragraph);
  const lineHeight = fontSize * (paragraph.lineSpacing ?? (paragraph.headingLevel ? 1.28 : 1.35));
  const baseIndent = (paragraph.indentLeft ?? 0) + (paragraph.firstLineIndent ?? 0) - (paragraph.hangingIndent ?? 0);
  const listIndent = paragraph.list ? 18 + paragraph.list.level * 18 : 0;
  const indent = Math.max(0, baseIndent + listIndent);
  const prefix = paragraph.list ? `${formatListPrefix(paragraph.list)} ` : "";
  const maxWidth = Math.max(80, document.pageSize.width - document.pageSize.marginLeft - document.pageSize.marginRight - indent - (paragraph.indentRight ?? 0));
  const font = pickFont(ctx, firstRun, Boolean(paragraph.headingLevel));
  const lines = wrapText(`${prefix}${text}`, maxWidth, font, fontSize);
  const spacingBefore = paragraph.spacingBefore ?? 0;
  const spacingAfter = paragraph.spacingAfter ?? 6;
  const blockHeight = spacingBefore + Math.max(lineHeight, lines.length * lineHeight) + spacingAfter;
  ensureSpace(ctx, document, blockHeight, Boolean(paragraph.keepWithNext || paragraph.keepLinesTogether || paragraph.headingLevel));
  const blockTop = ctx.y;
  ctx.y -= spacingBefore;
  if (paragraph.headingLevel && text.trim()) ctx.headings.push({ blockId: paragraph.id, title: text.trim(), level: paragraph.headingLevel, pageNumber: ctx.pageNumber });
  const color = colorToRgb(paragraph.runs.find((run) => run.color)?.color ?? (paragraph.runs.some((run) => run.hyperlink) ? "#2563eb" : "#111827"));
  for (const line of lines) {
    const lineX = getAlignedX(ctx, document, paragraph, line, font, fontSize, indent);
    const lineY = ctx.y - lineHeight;
    const highlighted = paragraph.runs.find((run) => run.backgroundColor);
    if (highlighted?.backgroundColor) {
      ctx.page.drawRectangle({ x: lineX, y: lineY - 1, width: font.widthOfTextAtSize(line, fontSize), height: lineHeight, color: colorToRgb(highlighted.backgroundColor), opacity: 0.72 });
    }
    const verticalRun = paragraph.runs.find((run) => run.verticalAlign && run.verticalAlign !== "baseline");
    const yOffset = verticalRun?.verticalAlign === "superscript" ? fontSize * 0.32 : verticalRun?.verticalAlign === "subscript" ? -fontSize * 0.22 : 0;
    ctx.page.drawText(line, { x: lineX, y: lineY + yOffset, size: verticalRun ? fontSize * 0.82 : fontSize, font, color });
    drawRunDecorations(ctx.page, line, lineX, lineY, font, fontSize, paragraph.runs);
    addLineLink(ctx, paragraph, line, lineX, lineY, font, fontSize, lineHeight);
    ctx.y -= lineHeight;
  }
  ctx.y -= spacingAfter;
  recordMapping(ctx, paragraph, document.pageSize.marginLeft + indent, ctx.y, maxWidth, blockTop - ctx.y);
}

async function drawImage(ctx: LayoutContext, document: DocxIntermediateDocument, block: DocxImageBlock) {
  const image = ctx.images.get(block.relationshipId);
  if (!image) return;
  image.used = true;
  if (image.mimeType !== "image/png" && image.mimeType !== "image/jpeg") {
    ctx.warnings.push({ code: "image-render-fallback", message: `${image.name} was preserved as a project asset but could not be rendered into first-phase PDF pages.` });
    recordMapping(ctx, block, document.pageSize.marginLeft, ctx.y - 32, 160, 32);
    return;
  }
  if (block.floating && ctx.fidelityMode !== "high") ctx.warnings.push({ code: "floating-image-simplified", message: `${image.name} uses floating Word positioning and was placed inline for ${ctx.fidelityMode} import.` });
  const maxWidth = document.pageSize.width - document.pageSize.marginLeft - document.pageSize.marginRight;
  const maxHeight = Math.min(ctx.fidelityMode === "high" ? 320 : 240, document.pageSize.height * 0.42);
  const naturalWidth = block.width || image.width;
  const naturalHeight = block.height || image.height;
  const scale = Math.min(maxWidth / Math.max(1, naturalWidth), maxHeight / Math.max(1, naturalHeight), ctx.fidelityMode === "fast" ? 0.85 : 1);
  const width = Math.max(24, naturalWidth * scale);
  const height = Math.max(24, naturalHeight * scale);
  ensureSpace(ctx, document, height + 12);
  const embedded = image.mimeType === "image/jpeg" ? await ctx.pdf.embedJpg(image.bytes) : await ctx.pdf.embedPng(image.bytes);
  ctx.page.drawImage(embedded, { x: document.pageSize.marginLeft, y: ctx.y - height, width, height });
  recordMapping(ctx, block, document.pageSize.marginLeft, ctx.y - height, width, height);
  ctx.y -= height + 12;
}

function drawTable(ctx: LayoutContext, document: DocxIntermediateDocument, table: DocxTableBlock) {
  const columnCount = Math.max(1, table.columnWidths?.length ?? 0, ...table.rows.map((row) => row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0)));
  const availableWidth = document.pageSize.width - document.pageSize.marginLeft - document.pageSize.marginRight;
  const rawWidths = table.columnWidths && table.columnWidths.length >= columnCount ? table.columnWidths.slice(0, columnCount) : Array.from({ length: columnCount }, () => availableWidth / columnCount);
  const rawTotal = rawWidths.reduce((sum, width) => sum + width, 0) || availableWidth;
  const scale = rawTotal > availableWidth ? availableWidth / rawTotal : 1;
  if (rawTotal > availableWidth) ctx.warnings.push({ code: "table-scaled", message: "A Word table exceeded the page width and was scaled to fit without clipping columns." });
  const columnWidths = rawWidths.map((width) => width * scale);
  const fontSize = ctx.fidelityMode === "fast" ? 8.5 : 9;
  const lineHeight = fontSize * 1.35;
  const tableTop = ctx.y;
  for (const [rowIndex, row] of table.rows.entries()) {
    const cellTexts = row.map((cell) => cell.blocks.map(getParagraphText).join("\n"));
    const rowLines = cellTexts.map((text, index) => wrapText(text || " ", Math.max(36, getCellWidth(columnWidths, row, index) - 8), ctx.fontRegular, fontSize));
    const rowHeight = Math.max(24, Math.max(...rowLines.map((lines) => lines.length)) * lineHeight + 10);
    ensureSpace(ctx, document, rowHeight, table.headerRows?.includes(rowIndex));
    const y = ctx.y - rowHeight;
    let x = document.pageSize.marginLeft;
    for (let index = 0; index < row.length; index += 1) {
      const cell = row[index];
      const width = getCellWidth(columnWidths, row, index);
      ctx.page.drawRectangle({ x, y, width, height: rowHeight, borderColor: rgb(0.72, 0.76, 0.82), borderWidth: 0.6, color: cell.background ? colorToRgb(cell.background) : undefined });
      const lines = rowLines[index] ?? [""];
      lines.slice(0, Math.floor((rowHeight - 8) / lineHeight)).forEach((line, lineIndex) => {
        ctx.page.drawText(line, { x: x + 4, y: ctx.y - 14 - lineIndex * lineHeight, size: fontSize, font: ctx.fontRegular, color: rgb(0.08, 0.1, 0.16) });
      });
      x += width;
    }
    ctx.y -= rowHeight;
  }
  recordMapping(ctx, table, document.pageSize.marginLeft, ctx.y, availableWidth, tableTop - ctx.y);
  ctx.y -= 10;
}

function getCellWidth(columnWidths: number[], row: Array<{ gridSpan?: number }>, cellIndex: number) {
  const start = row.slice(0, cellIndex).reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0);
  const span = row[cellIndex]?.gridSpan ?? 1;
  return columnWidths.slice(start, start + span).reduce((sum, width) => sum + width, 0) || columnWidths[start] || 72;
}

function addLineLink(ctx: LayoutContext, paragraph: DocxParagraphBlock, line: string, x: number, y: number, font: PDFFont, fontSize: number, lineHeight: number) {
  const linkRun = paragraph.runs.find((run) => run.hyperlink && isExternalLink(run.hyperlink));
  if (!linkRun?.hyperlink) return;
  const width = font.widthOfTextAtSize(line, fontSize);
  ctx.links.push({ blockId: paragraph.id, text: line.trim(), url: normalizeExternalLink(linkRun.hyperlink), pageNumber: ctx.pageNumber, x, y: y - 2, width, height: lineHeight });
}

function addExternalLinkAnnotations(pdf: PDFDocument, links: DocxRenderedLink[]) {
  const pages = pdf.getPages();
  for (const link of links) {
    const page = pages[link.pageNumber - 1];
    if (!page) continue;
    const context = pdf.context;
    const action = context.obj({ Type: "Action", S: "URI", URI: PDFHexString.fromText(link.url) });
    const annotation = context.register(
      context.obj({
        Type: "Annot",
        Subtype: "Link",
        Rect: [link.x, link.y, link.x + link.width, link.y + link.height],
        Border: [0, 0, 0],
        A: action,
      })
    );
    const existing = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
    const annots = existing ?? context.obj([]);
    annots.push(annotation);
    page.node.set(PDFName.of("Annots"), annots);
  }
}

function recordMapping(ctx: LayoutContext, block: DocxBlock, x: number, y: number, width: number, height: number) {
  ctx.sourceMappings.push({ blockId: block.id, kind: block.type, sourcePath: block.sourcePath, sourceOrder: block.sourceOrder, pageNumber: ctx.pageNumber, x, y, width, height });
}

function getParagraphText(paragraph: DocxParagraphBlock) {
  return paragraph.runs.map((run) => run.text).join("").replace(/\s+\n/g, "\n").replace(/\n\s+/g, "\n");
}

function getParagraphFontSize(paragraph: DocxParagraphBlock) {
  if (paragraph.headingLevel === 1) return 22;
  if (paragraph.headingLevel === 2) return 17;
  if (paragraph.headingLevel === 3) return 14;
  return paragraph.runs.find((run) => run.fontSize)?.fontSize ?? 11;
}

function pickFont(ctx: LayoutContext, run?: DocxTextRun, forceBold = false) {
  const bold = forceBold || run?.bold;
  if (bold && run?.italic) return ctx.fontBoldItalic;
  if (bold) return ctx.fontBold;
  if (run?.italic) return ctx.fontItalic;
  return ctx.fontRegular;
}

function getAlignedX(ctx: LayoutContext, document: DocxIntermediateDocument, paragraph: DocxParagraphBlock, line: string, font: PDFFont, fontSize: number, indent: number) {
  const left = document.pageSize.marginLeft + indent;
  if (paragraph.alignment === "center") return document.pageSize.width / 2 - font.widthOfTextAtSize(line, fontSize) / 2;
  if (paragraph.alignment === "right") return document.pageSize.width - document.pageSize.marginRight - font.widthOfTextAtSize(line, fontSize);
  return left;
}

function wrapText(text: string, maxWidth: number, font: PDFFont, fontSize: number) {
  const rawLines = text.split("\n");
  const lines: string[] = [];
  for (const rawLine of rawLines) {
    const words = rawLine.split(/\s+/).filter(Boolean);
    let current = "";
    for (const word of words) {
      const candidate = current ? `${current} ${word}` : word;
      if (font.widthOfTextAtSize(candidate, fontSize) <= maxWidth || !current) current = candidate;
      else {
        lines.push(current);
        current = word;
      }
    }
    lines.push(current || " ");
  }
  return lines;
}

function drawRunDecorations(page: PDFPage, line: string, x: number, y: number, font: PDFFont, fontSize: number, runs: DocxTextRun[]) {
  const hasUnderline = runs.some((run) => run.underline || run.hyperlink);
  const hasStrike = runs.some((run) => run.strike);
  const width = font.widthOfTextAtSize(line, fontSize);
  if (hasUnderline) page.drawLine({ start: { x, y: y - 1.5 }, end: { x: x + width, y: y - 1.5 }, thickness: 0.5, color: rgb(0.1, 0.32, 0.72) });
  if (hasStrike) page.drawLine({ start: { x, y: y + fontSize * 0.35 }, end: { x: x + width, y: y + fontSize * 0.35 }, thickness: 0.5, color: rgb(0.08, 0.1, 0.16) });
}

function formatListPrefix(list: NonNullable<DocxParagraphBlock["list"]>) {
  if (list.kind === "bullet") return list.text?.replace(/%[0-9]/g, "") || "\u2022";
  const value = list.kind === "romanLower" ? toRoman(list.index).toLowerCase() : list.kind === "romanUpper" ? toRoman(list.index) : list.kind === "alphaLower" ? toAlpha(list.index).toLowerCase() : list.kind === "alphaUpper" ? toAlpha(list.index) : String(list.index);
  return list.text?.replace(/%[0-9]/g, value) ?? `${value}.`;
}

function toRoman(value: number) {
  const numerals: Array<[number, string]> = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let remaining = Math.max(1, Math.floor(value));
  let output = "";
  for (const [amount, numeral] of numerals) {
    while (remaining >= amount) {
      output += numeral;
      remaining -= amount;
    }
  }
  return output;
}

function toAlpha(value: number) {
  let remaining = Math.max(1, Math.floor(value));
  let output = "";
  while (remaining > 0) {
    remaining -= 1;
    output = String.fromCharCode(65 + (remaining % 26)) + output;
    remaining = Math.floor(remaining / 26);
  }
  return output;
}

function isExternalLink(url: string) {
  return /^https?:\/\//i.test(url) || /^mailto:/i.test(url);
}

function normalizeExternalLink(url: string) {
  return /^[a-z][a-z0-9+.-]*:/i.test(url) ? url : `https://${url}`;
}

function colorToRgb(hex: string) {
  const normalized = hex.replace("#", "");
  const value = Number.parseInt(normalized.length === 6 ? normalized : "111827", 16);
  return rgb(((value >> 16) & 255) / 255, ((value >> 8) & 255) / 255, (value & 255) / 255);
}
