import { PDFDocument, StandardFonts, rgb, type PDFFont, type PDFPage } from "pdf-lib";
import type { DocxBlock, DocxExtractedImage, DocxHeading, DocxImageBlock, DocxImportOptions, DocxIntermediateDocument, DocxParagraphBlock, DocxTableBlock, DocxTextRun } from "./types";

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
  warnings: DocxIntermediateDocument["warnings"];
};

export async function renderDocxToPdf(document: DocxIntermediateDocument, fallbackFont: DocxImportOptions["fallbackFont"] = "Helvetica") {
  const pdf = await PDFDocument.create();
  const fontSet = getFontSet(fallbackFont);
  const fontRegular = await pdf.embedFont(fontSet.regular);
  const fontBold = await pdf.embedFont(fontSet.bold);
  const fontItalic = await pdf.embedFont(fontSet.italic);
  const fontBoldItalic = await pdf.embedFont(fontSet.boldItalic);
  const images = new Map(document.images.map((image) => [image.relationshipId, image]));
  const headings: DocxHeading[] = [];
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
    headings,
    warnings: document.warnings,
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

  return { bytes: await pdf.save(), pageCount: pdf.getPageCount(), headings };
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

function ensureSpace(ctx: LayoutContext, document: DocxIntermediateDocument, height: number) {
  if (ctx.y - height < document.pageSize.marginBottom) addPage(ctx, document);
}

function drawParagraph(ctx: LayoutContext, document: DocxIntermediateDocument, paragraph: DocxParagraphBlock) {
  const text = getParagraphText(paragraph);
  if (!text.trim() && !paragraph.headingLevel) return;
  const fontSize = getParagraphFontSize(paragraph);
  const lineHeight = fontSize * (paragraph.headingLevel ? 1.28 : 1.35);
  const indent = (paragraph.indentLeft ?? 0) + (paragraph.list ? 18 + paragraph.list.level * 18 : 0);
  const prefix = paragraph.list ? `${paragraph.list.kind === "bullet" ? "\u2022" : `${paragraph.list.index}.`} ` : "";
  const maxWidth = document.pageSize.width - document.pageSize.marginLeft - document.pageSize.marginRight - indent;
  const lines = wrapText(`${prefix}${text}`, maxWidth, pickFont(ctx, paragraph.runs[0]), fontSize);
  const blockHeight = (paragraph.spacingBefore ?? 0) + Math.max(lineHeight, lines.length * lineHeight) + (paragraph.spacingAfter ?? 6);
  ensureSpace(ctx, document, blockHeight);
  ctx.y -= paragraph.spacingBefore ?? 0;
  if (paragraph.headingLevel && text.trim()) {
    ctx.headings.push({ title: text.trim(), level: paragraph.headingLevel, pageNumber: ctx.pageNumber });
  }
  const font = pickFont(ctx, paragraph.runs[0], Boolean(paragraph.headingLevel));
  const color = colorToRgb(paragraph.runs.find((run) => run.color)?.color ?? "#111827");
  for (const line of lines) {
    const x = getAlignedX(ctx, document, paragraph, line, font, fontSize, indent);
    ctx.page.drawText(line, { x, y: ctx.y - lineHeight, size: fontSize, font, color });
    drawRunDecorations(ctx.page, line, x, ctx.y - lineHeight, font, fontSize, paragraph.runs);
    ctx.y -= lineHeight;
  }
  ctx.y -= paragraph.spacingAfter ?? 6;
}

async function drawImage(ctx: LayoutContext, document: DocxIntermediateDocument, block: DocxImageBlock) {
  const image = ctx.images.get(block.relationshipId);
  if (!image) return;
  image.used = true;
  if (image.mimeType !== "image/png" && image.mimeType !== "image/jpeg") {
    ctx.warnings.push({ code: "image-render-fallback", message: `${image.name} was preserved as a project asset but could not be rendered into first-phase PDF pages.` });
    return;
  }
  const maxWidth = document.pageSize.width - document.pageSize.marginLeft - document.pageSize.marginRight;
  const maxHeight = Math.min(240, document.pageSize.height * 0.35);
  const scale = Math.min(maxWidth / Math.max(1, image.width), maxHeight / Math.max(1, image.height), 1);
  const width = Math.max(24, image.width * scale);
  const height = Math.max(24, image.height * scale);
  ensureSpace(ctx, document, height + 12);
  const embedded = image.mimeType === "image/jpeg" ? await ctx.pdf.embedJpg(image.bytes) : await ctx.pdf.embedPng(image.bytes);
  ctx.page.drawImage(embedded, { x: document.pageSize.marginLeft, y: ctx.y - height, width, height });
  ctx.y -= height + 12;
}

function drawTable(ctx: LayoutContext, document: DocxIntermediateDocument, table: DocxTableBlock) {
  const columnCount = Math.max(1, ...table.rows.map((row) => row.length));
  const tableWidth = document.pageSize.width - document.pageSize.marginLeft - document.pageSize.marginRight;
  const columnWidth = tableWidth / columnCount;
  const fontSize = 9;
  const lineHeight = 12;
  for (const row of table.rows) {
    const cellTexts = row.map((cell) => cell.blocks.map(getParagraphText).join("\n"));
    const rowLines = cellTexts.map((text) => wrapText(text || " ", columnWidth - 8, ctx.fontRegular, fontSize));
    const rowHeight = Math.max(24, Math.max(...rowLines.map((lines) => lines.length)) * lineHeight + 10);
    ensureSpace(ctx, document, rowHeight);
    const y = ctx.y - rowHeight;
    for (let index = 0; index < columnCount; index += 1) {
      const x = document.pageSize.marginLeft + index * columnWidth;
      const background = row[index]?.background;
      ctx.page.drawRectangle({ x, y, width: columnWidth, height: rowHeight, borderColor: rgb(0.75, 0.78, 0.84), borderWidth: 0.6, color: background ? colorToRgb(background) : undefined });
      const lines = rowLines[index] ?? [""];
      lines.slice(0, Math.floor((rowHeight - 8) / lineHeight)).forEach((line, lineIndex) => {
        ctx.page.drawText(line, { x: x + 4, y: ctx.y - 14 - lineIndex * lineHeight, size: fontSize, font: ctx.fontRegular, color: rgb(0.08, 0.1, 0.16) });
      });
    }
    ctx.y -= rowHeight;
  }
  ctx.y -= 10;
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

function colorToRgb(hex: string) {
  const normalized = hex.replace("#", "");
  const value = Number.parseInt(normalized.length === 6 ? normalized : "111827", 16);
  return rgb(((value >> 16) & 255) / 255, ((value >> 8) & 255) / 255, (value & 255) / 255);
}
