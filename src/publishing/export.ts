import { degrees, PDFDocument, rgb, StandardFonts, type PDFFont, type PDFImage } from "pdf-lib";
import { getPublishingPreviewItems } from "./layout";
import type { ResolvedPageNumber } from "./numberingResolver";
import type { PublishingPageLike, PublishingSettings, PublishingTextStyle, PublishingTokenContext } from "./types";

type ImageAsset = {
  id: string;
  dataUrl?: string;
  mimeType?: string;
};

export async function drawPublishingMarksToPdf({
  pdf,
  pages,
  settings,
  tokenContext,
  currentPage,
  selectedPageIds,
  imageAssets,
  resolvedPageNumbers,
  layer,
}: {
  pdf: PDFDocument;
  pages: PublishingPageLike[];
  settings: PublishingSettings;
  tokenContext: Omit<PublishingTokenContext, "pageNumber" | "pageCount" | "pageLabel">;
  currentPage: number;
  selectedPageIds: string[];
  imageAssets: ImageAsset[];
  resolvedPageNumbers?: Map<string, ResolvedPageNumber>;
  layer: "watermark" | "foreground";
}) {
  const fontCache = new Map<string, PDFFont>();
  const imageCache = new Map<string, PDFImage>();

  for (const [index, pageView] of pages.entries()) {
    const page = pdf.getPage(index);
    const pdfSize = page.getSize();
    const scaleX = pdfSize.width / pageView.width;
    const scaleY = pdfSize.height / pageView.height;
    const previewItems = getPublishingPreviewItems({
      settings,
      page: pageView,
      pages,
      currentPage,
      selectedPageIds,
      tokenContext,
      imageAssetIds: imageAssets.map((asset) => asset.id),
      resolvedPageNumbers,
    });

    for (const item of previewItems) {
      if (item.kind === "safeArea") continue;
      if (layer === "watermark" && item.kind !== "watermark") continue;
      if (layer === "foreground" && item.kind === "watermark") continue;
      if (item.kind === "missingImage") continue;

      if (item.kind === "watermark" && settings.watermark.type === "image" && settings.watermark.imageAssetId) {
        const asset = imageAssets.find((candidate) => candidate.id === settings.watermark.imageAssetId);
        if (asset?.dataUrl) {
          const embedded = await getEmbeddedImage(pdf, imageCache, asset);
          if (embedded) {
            const width = Math.min(pdfSize.width * 0.72, embedded.width * settings.watermark.scale * 0.35);
            const height = width * (embedded.height / embedded.width);
            const x = item.x * scaleX - width / 2;
            const y = pdfSize.height - item.y * scaleY - height / 2;
            page.drawImage(embedded, {
              x,
              y,
              width,
              height,
              rotate: degrees(settings.watermark.rotation),
              opacity: settings.watermark.opacity,
            });
          }
        }
        continue;
      }

      if (item.kind === "image" && item.assetId) {
        const asset = imageAssets.find((candidate) => candidate.id === item.assetId);
        if (asset?.dataUrl) {
          const embedded = await getEmbeddedImage(pdf, imageCache, asset);
          if (embedded && item.width && item.height) {
            const width = item.width * scaleX;
            const height = item.height * scaleY;
            page.drawImage(embedded, {
              x: item.align === "center" ? item.x * scaleX - width / 2 : item.align === "right" ? item.x * scaleX - width : item.x * scaleX,
              y: pdfSize.height - item.y * scaleY - height / 2,
              width,
              height,
              opacity: item.opacity ?? 1,
            });
          }
        }
        continue;
      }

      const style = item.style;
      if (!item.text || !style) continue;
      const font = await getFont(pdf, fontCache, style);
      const fontSize = style.fontSize * (item.kind === "watermark" ? settings.watermark.scale : 1);
      const textWidth = font.widthOfTextAtSize(item.text, fontSize);
      const x = item.align === "center" ? item.x * scaleX - textWidth / 2 : item.align === "right" ? item.x * scaleX - textWidth : item.x * scaleX;
      const y = pdfSize.height - item.y * scaleY - fontSize * 0.35;
      page.drawText(item.text, {
        x,
        y,
        size: fontSize,
        font,
        color: hexToRgb(style.color),
        opacity: item.opacity ?? style.opacity,
        rotate: item.rotation ? degrees(item.rotation) : undefined,
      });
    }
  }
}

async function getFont(pdf: PDFDocument, cache: Map<string, PDFFont>, style: PublishingTextStyle) {
  const name = getStandardFont(style);
  const existing = cache.get(name);
  if (existing) return existing;
  const font = await pdf.embedFont(name);
  cache.set(name, font);
  return font;
}

function getStandardFont(style: PublishingTextStyle) {
  if (style.fontFamily === "Times Roman") {
    if (style.bold && style.italic) return StandardFonts.TimesRomanBoldItalic;
    if (style.bold) return StandardFonts.TimesRomanBold;
    if (style.italic) return StandardFonts.TimesRomanItalic;
    return StandardFonts.TimesRoman;
  }
  if (style.fontFamily === "Courier") {
    if (style.bold && style.italic) return StandardFonts.CourierBoldOblique;
    if (style.bold) return StandardFonts.CourierBold;
    if (style.italic) return StandardFonts.CourierOblique;
    return StandardFonts.Courier;
  }
  if (style.bold && style.italic) return StandardFonts.HelveticaBoldOblique;
  if (style.bold) return StandardFonts.HelveticaBold;
  if (style.italic) return StandardFonts.HelveticaOblique;
  return StandardFonts.Helvetica;
}

async function getEmbeddedImage(pdf: PDFDocument, cache: Map<string, PDFImage>, asset: ImageAsset) {
  const existing = cache.get(asset.id);
  if (existing) return existing;
  if (!asset.dataUrl) return null;
  const bytes = dataUrlToBytes(asset.dataUrl);
  const embedded = asset.mimeType === "image/jpeg" ? await pdf.embedJpg(bytes) : await pdf.embedPng(bytes);
  cache.set(asset.id, embedded);
  return embedded;
}

function dataUrlToBytes(dataUrl: string) {
  const [, encoded = ""] = dataUrl.split(",", 2);
  const binary = atob(encoded);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

function hexToRgb(hex: string) {
  const normalized = hex.replace("#", "");
  const full = normalized.length === 3 ? normalized.split("").map((char) => char + char).join("") : normalized.padEnd(6, "0").slice(0, 6);
  return rgb(parseInt(full.slice(0, 2), 16) / 255, parseInt(full.slice(2, 4), 16) / 255, parseInt(full.slice(4, 6), 16) / 255);
}
