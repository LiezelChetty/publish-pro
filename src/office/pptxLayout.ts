import { PDFArray, PDFHexString, PDFName, PDFNumber, PDFDocument, StandardFonts, rgb, type PDFFont, type PDFPage } from "pdf-lib";
import type { PptxHyperlink, PptxImageAsset, PptxIntermediatePresentation, PptxSlide, PptxSlideElement, PptxSourceMapping, PptxTextRun } from "./pptxTypes";

type PptxRenderResult = {
  bytes: Uint8Array;
  pageCount: number;
  sourceMappings: PptxSourceMapping[];
  links: PptxHyperlink[];
};

export async function renderPptxToPdf(presentation: PptxIntermediatePresentation): Promise<PptxRenderResult> {
  const pdf = await PDFDocument.create();
  const fontRegular = await pdf.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const imageMap = new Map(presentation.imageAssets.map((asset) => [asset.id, asset]));
  const imageCache = new Map<string, Awaited<ReturnType<typeof pdf.embedPng>>>();
  const sourceMappings: PptxSourceMapping[] = [];
  const links: PptxHyperlink[] = [];

  for (const slide of presentation.slides) {
    const page = pdf.addPage([presentation.slideSize.width, presentation.slideSize.height]);
    drawSlideBackground(page, slide, presentation);
    sourceMappings.push({ slideId: slide.id, elementId: slide.id, kind: "slide", sourcePath: slide.sourcePath, pageNumber: slide.slideIndex, x: 0, y: 0, width: presentation.slideSize.width, height: presentation.slideSize.height });
    for (const element of slide.elements) {
      if (element.kind === "text") drawTextBlock(page, element, fontRegular, fontBold, presentation.slideSize.height);
      if (element.kind === "shape") drawShape(page, element, presentation.slideSize.height);
      if (element.kind === "table") drawTable(page, element, fontRegular, presentation.slideSize.height);
      if (element.kind === "chart") drawChartPlaceholder(page, element, fontBold, presentation.slideSize.height);
      if (element.kind === "image") await drawImage(pdf, page, element, imageMap, imageCache, presentation.slideSize.height, presentation);
      sourceMappings.push({ slideId: slide.id, elementId: element.id, kind: element.kind, sourcePath: element.sourcePath, pageNumber: slide.slideIndex, x: element.x, y: element.y, width: element.width, height: element.height });
    }
    links.push(...slide.hyperlinks);
  }

  addLinkAnnotations(pdf, links);
  return { bytes: await pdf.save(), pageCount: pdf.getPageCount(), sourceMappings, links };
}

function drawSlideBackground(page: PDFPage, slide: PptxSlide, presentation: PptxIntermediatePresentation) {
  page.drawRectangle({ x: 0, y: 0, width: presentation.slideSize.width, height: presentation.slideSize.height, color: colorToRgb(slide.background ?? "#ffffff") });
}

function drawTextBlock(page: PDFPage, element: Extract<PptxSlideElement, { kind: "text" }>, fontRegular: PDFFont, fontBold: PDFFont, slideHeight: number) {
  const padding = 4;
  let y = slideHeight - element.y - padding - 14;
  for (const paragraph of element.paragraphs) {
    const lines = wrapRuns(paragraph, Math.max(24, element.width - padding * 2), fontRegular);
    for (const line of lines) {
      const firstRun = paragraph[0];
      const fontSize = firstRun?.fontSize ?? 14;
      const font = firstRun?.bold ? fontBold : fontRegular;
      page.drawText(line, { x: element.x + padding, y, size: fontSize, font, color: colorToRgb(firstRun?.color ?? "#111827") });
      if (firstRun?.underline) page.drawLine({ start: { x: element.x + padding, y: y - 2 }, end: { x: element.x + padding + font.widthOfTextAtSize(line, fontSize), y: y - 2 }, thickness: 0.6, color: colorToRgb(firstRun?.color ?? "#111827") });
      if (firstRun?.strike) page.drawLine({ start: { x: element.x + padding, y: y + fontSize * 0.35 }, end: { x: element.x + padding + font.widthOfTextAtSize(line, fontSize), y: y + fontSize * 0.35 }, thickness: 0.6, color: colorToRgb(firstRun?.color ?? "#111827") });
      y -= fontSize * 1.25;
      if (y < slideHeight - element.y - element.height) return;
    }
  }
}

function drawShape(page: PDFPage, element: Extract<PptxSlideElement, { kind: "shape" }>, slideHeight: number) {
  const y = slideHeight - element.y - element.height;
  const options = { x: element.x, y, width: element.width, height: element.height, color: colorToRgb(element.fill ?? "#f8fafc"), borderColor: colorToRgb(element.stroke ?? "#94a3b8"), borderWidth: element.strokeWidth ?? 1 };
  if (element.preset.includes("ellipse")) page.drawEllipse({ x: element.x + element.width / 2, y: y + element.height / 2, xScale: element.width / 2, yScale: element.height / 2, color: options.color, borderColor: options.borderColor, borderWidth: options.borderWidth });
  else page.drawRectangle(options);
}

function drawTable(page: PDFPage, element: Extract<PptxSlideElement, { kind: "table" }>, font: PDFFont, slideHeight: number) {
  const rows = element.rows.length || 1;
  const columns = Math.max(1, ...element.rows.map((row) => row.length));
  const cellWidth = element.width / columns;
  const cellHeight = element.height / rows;
  const top = slideHeight - element.y;
  for (let row = 0; row < rows; row += 1) {
    for (let column = 0; column < columns; column += 1) {
      const x = element.x + column * cellWidth;
      const y = top - (row + 1) * cellHeight;
      page.drawRectangle({ x, y, width: cellWidth, height: cellHeight, borderColor: rgb(0.65, 0.7, 0.78), borderWidth: 0.6, color: row === 0 ? rgb(0.94, 0.96, 0.98) : undefined });
      const text = element.rows[row]?.[column] ?? "";
      if (text) page.drawText(text.slice(0, 60), { x: x + 4, y: y + Math.max(4, cellHeight - 14), size: 9, font, color: rgb(0.08, 0.1, 0.16) });
    }
  }
}

function drawChartPlaceholder(page: PDFPage, element: Extract<PptxSlideElement, { kind: "chart" }>, font: PDFFont, slideHeight: number) {
  const y = slideHeight - element.y - element.height;
  page.drawRectangle({ x: element.x, y, width: element.width, height: element.height, color: rgb(0.96, 0.97, 0.99), borderColor: rgb(0.55, 0.6, 0.7), borderWidth: 1 });
  page.drawText("Chart preview", { x: element.x + 16, y: y + element.height / 2, size: 14, font, color: rgb(0.2, 0.25, 0.34) });
}

async function drawImage(pdf: PDFDocument, page: PDFPage, element: Extract<PptxSlideElement, { kind: "image" }>, imageMap: Map<string, PptxImageAsset>, imageCache: Map<string, Awaited<ReturnType<typeof pdf.embedPng>>>, slideHeight: number, presentation: PptxIntermediatePresentation) {
  const asset = element.assetId ? imageMap.get(element.assetId) : undefined;
  if (!asset) return;
  if (asset.mimeType !== "image/png" && asset.mimeType !== "image/jpeg") {
    presentation.warnings.push({ code: "pptx-image-fallback", category: "images", message: `${asset.name} is preserved as a Project Asset but was not embedded into the generated PDF in this phase.` });
    return;
  }
  let embedded = imageCache.get(asset.id);
  if (!embedded) {
    embedded = asset.mimeType === "image/jpeg" ? await pdf.embedJpg(asset.bytes) : await pdf.embedPng(asset.bytes);
    imageCache.set(asset.id, embedded);
  }
  page.drawImage(embedded, { x: element.x, y: slideHeight - element.y - element.height, width: element.width, height: element.height });
  asset.used = true;
}

function addLinkAnnotations(pdf: PDFDocument, links: PptxHyperlink[]) {
  const pages = pdf.getPages();
  for (const link of links) {
    const page = pages[link.slideIndex - 1];
    if (!page || !link.url) continue;
    const context = pdf.context;
    const action = context.obj({ Type: "Action", S: "URI", URI: PDFHexString.fromText(link.url) });
    const annotation = context.register(context.obj({ Type: "Annot", Subtype: "Link", Rect: [link.x, page.getHeight() - link.y - link.height, link.x + link.width, page.getHeight() - link.y], Border: [0, 0, 0], A: action }));
    const existing = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
    const annots = existing ?? context.obj([]);
    annots.push(annotation);
    page.node.set(PDFName.of("Annots"), annots);
  }
}

function wrapRuns(runs: PptxTextRun[], maxWidth: number, font: PDFFont) {
  const text = runs.map((run) => run.text).join("");
  const fontSize = runs[0]?.fontSize ?? 14;
  const words = text.split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let current = "";
  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word;
    if (font.widthOfTextAtSize(candidate, fontSize) <= maxWidth || !current) current = candidate;
    else {
      lines.push(current);
      current = word;
    }
  }
  if (current) lines.push(current);
  return lines.length ? lines : [text];
}

function colorToRgb(hex: string) {
  const normalized = hex.replace("#", "");
  const value = Number.parseInt(normalized.length === 6 ? normalized : "111827", 16);
  return rgb(((value >> 16) & 255) / 255, ((value >> 8) & 255) / 255, (value & 255) / 255);
}
