import { resolvePublishingPageIds } from "./pageRanges";
import { formatPageNumber, renderPublishingTokens } from "./tokens";
import type { ResolvedPageNumber } from "./numberingResolver";
import type { HeaderFooterZone, HeaderFooterZoneImage, PublishingPageLike, PublishingPreviewItem, PublishingSettings, PublishingTokenContext, PublishingUnit, PublishingZone } from "./types";

export function unitToPoints(value: number, unit: PublishingUnit) {
  if (unit === "mm") return (value / 25.4) * 72;
  if (unit === "cm") return (value / 2.54) * 72;
  if (unit === "in") return value * 72;
  return value;
}

export function getPublishingPreviewItems({
  settings,
  page,
  pages,
  currentPage,
  selectedPageIds,
  tokenContext,
  imageAssetIds = [],
  resolvedPageNumbers,
}: {
  settings: PublishingSettings;
  page: PublishingPageLike;
  pages: PublishingPageLike[];
  currentPage: number;
  selectedPageIds: string[];
  tokenContext: Omit<PublishingTokenContext, "pageNumber" | "pageCount" | "pageLabel">;
  imageAssetIds?: string[];
  resolvedPageNumbers?: Map<string, ResolvedPageNumber>;
}): PublishingPreviewItem[] {
  const items: PublishingPreviewItem[] = [];
  const context = {
    ...tokenContext,
    pageNumber: page.pageNumber,
    pageCount: pages.length,
    pageLabel: page.label,
  };

  if (settings.watermark.enabled && resolvePublishingPageIds(settings.watermark.target, pages, currentPage, selectedPageIds).has(page.id)) {
    const position = getWatermarkPosition(settings.watermark.position, page, settings.watermark.customX, settings.watermark.customY, settings.watermark.unit);
    const item = {
      id: `watermark-${page.id}`,
      kind: "watermark",
      text: settings.watermark.type === "text" ? renderPublishingTokens(settings.watermark.text, context) : "Image watermark",
      x: position.x,
      y: position.y,
      align: "center",
      rotation: settings.watermark.rotation,
      style: settings.watermark.style,
      opacity: settings.watermark.opacity,
    } satisfies PublishingPreviewItem;
    if (settings.watermark.tiled) {
      const gap = Math.max(96, settings.watermark.style.fontSize * settings.watermark.scale * 2.8);
      let tileIndex = 0;
      for (let y = gap / 2; y < page.height; y += gap) {
        for (let x = gap / 2; x < page.width; x += gap) {
          items.push({ ...item, id: `${item.id}-tile-${tileIndex}`, x, y });
          tileIndex += 1;
        }
      }
    } else {
      items.push(item);
    }
  }

  if (settings.headerFooter.enabled && resolvePublishingPageIds(settings.headerFooter.target, pages, currentPage, selectedPageIds).has(page.id)) {
    const margin = unitToPoints(settings.headerFooter.margin, settings.headerFooter.unit);
    items.push({ id: `hf-safe-${page.id}`, kind: "safeArea", x: margin, y: margin, width: page.width - margin * 2, height: page.height - margin * 2 });
    for (const zone of ["headerLeft", "headerCenter", "headerRight", "footerLeft", "footerCenter", "footerRight"] as PublishingZone[]) {
      const zoneValue = getHeaderFooterZone(settings, zone, page.pageNumber);
      const zoneItems = getHeaderFooterZonePreviewItems({
        zone,
        zoneValue,
        page,
        pageId: page.id,
        margin,
        context,
        style: settings.headerFooter.style,
        imageAssetIds,
      });
      items.push(...zoneItems);
    }
  }

  if (settings.pageNumbers.enabled && resolvePublishingPageIds(settings.pageNumbers.target, pages, currentPage, selectedPageIds).has(page.id)) {
    const resolved = resolvedPageNumbers?.get(page.id);
    const pageValue = resolved?.value ?? settings.pageNumbers.startNumber + page.pageNumber - 1 + settings.pageNumbers.offset;
    const formatted = resolved?.display ?? formatPageNumber(settings.pageNumbers.format, pageValue, pages.length, settings.pageNumbers.customTemplate);
    const template = resolved ? formatted : settings.pageNumbers.format === "custom" ? formatted : `${settings.pageNumbers.prefix}${formatted}${settings.pageNumbers.suffix}`;
    if (resolved && !resolved.included) return items;
    const margin = unitToPoints(settings.pageNumbers.distanceFromEdge, settings.pageNumbers.unit);
    const offset = unitToPoints(settings.pageNumbers.horizontalOffset, settings.pageNumbers.unit);
    items.push({
      id: `page-number-${page.id}`,
      kind: "text",
      text: renderPublishingTokens(template, { ...context, pageNumber: pageValue }),
      ...getZonePosition(settings.pageNumbers.zone, page, margin, offset),
      style: settings.pageNumbers.style,
    });
  }

  return items;
}

export function getHeaderFooterZoneText(settings: PublishingSettings, zone: PublishingZone, pageNumber: number) {
  return getHeaderFooterZone(settings, zone, pageNumber).text;
}

export function getHeaderFooterZone(settings: PublishingSettings, zone: PublishingZone, pageNumber: number): HeaderFooterZone {
  const override = pageNumber === 1 ? settings.headerFooter.firstPage?.[zone] : pageNumber % 2 === 0 ? settings.headerFooter.evenPage?.[zone] : settings.headerFooter.oddPage?.[zone];
  if (settings.headerFooter.advanced && override !== undefined) {
    if (typeof override === "string" && override.trim() !== "") return { text: override };
    if (typeof override === "object" && (override.text.trim() !== "" || override.image?.assetId)) return override;
  }
  const area = zone.startsWith("header") ? settings.headerFooter.header : settings.headerFooter.footer;
  const side = zone.endsWith("Left") ? "left" : zone.endsWith("Center") ? "center" : "right";
  return area[side];
}

export function getZonePosition(zone: PublishingZone, page: PublishingPageLike, margin: number, horizontalOffset: number): { x: number; y: number; align: "left" | "center" | "right" } {
  const isHeader = zone.startsWith("header");
  const align = zone.endsWith("Left") ? "left" : zone.endsWith("Center") ? "center" : "right";
  const x = align === "left" ? margin + horizontalOffset : align === "center" ? page.width / 2 + horizontalOffset : page.width - margin + horizontalOffset;
  return { x, y: isHeader ? margin : page.height - margin, align };
}

function getWatermarkPosition(position: string, page: PublishingPageLike, customX: number, customY: number, unit: PublishingUnit) {
  const margin = unitToPoints(24, "mm");
  const custom = { x: unitToPoints(customX, unit), y: unitToPoints(customY, unit) };
  if (position === "custom") return custom;
  if (position === "topLeft") return { x: margin, y: margin };
  if (position === "topCenter") return { x: page.width / 2, y: margin };
  if (position === "topRight") return { x: page.width - margin, y: margin };
  if (position === "bottomLeft") return { x: margin, y: page.height - margin };
  if (position === "bottomCenter") return { x: page.width / 2, y: page.height - margin };
  if (position === "bottomRight") return { x: page.width - margin, y: page.height - margin };
  return { x: page.width / 2, y: page.height / 2 };
}

function getHeaderFooterZonePreviewItems({
  zone,
  zoneValue,
  page,
  pageId,
  margin,
  context,
  style,
  imageAssetIds,
}: {
  zone: PublishingZone;
  zoneValue: HeaderFooterZone;
  page: PublishingPageLike;
  pageId: string;
  margin: number;
  context: PublishingTokenContext;
  style: PublishingSettings["headerFooter"]["style"];
  imageAssetIds: string[];
}) {
  const items: PublishingPreviewItem[] = [];
  const image = zoneValue.image;
  const hasImage = Boolean(image?.assetId);
  const text = renderPublishingTokens(zoneValue.text ?? "", context);
  const hasText = text.trim().length > 0;
  if (!hasImage && !hasText) return items;

  const position = getZonePosition(zone, page, margin, 0);
  const lineHeight = style.fontSize * 1.25;
  const imageBox = image ? getClampedImageBox(image, page, margin) : null;
  const gap = image?.padding ?? 6;
  const layout = image?.layout ?? "imageBeforeText";
  const assetMissing = Boolean(image?.assetId && !imageAssetIds.includes(image.assetId));
  const baseId = `${zone}-${pageId}`;

  if (hasImage && image && imageBox) {
    const imageOnly = layout === "imageOnly" || !hasText;
    const imagePoint = getImagePoint(position, imageBox, layout, gap, lineHeight, imageOnly, image);
    const x = clamp(imagePoint.x + unitless(image.offsetX), margin + imageBox.width / 2, page.width - margin);
    const y = clamp(imagePoint.y + unitless(image.offsetY), margin + imageBox.height / 2, page.height - margin - imageBox.height / 2);
    items.push({
      id: `${baseId}-image`,
      kind: assetMissing ? "missingImage" : "image",
      assetId: image.assetId,
      text: assetMissing ? "Missing image" : undefined,
      x,
      y,
      width: imageBox.width,
      height: imageBox.height,
      align: image.horizontalAlign,
      opacity: image.opacity,
    });
  }

  if (hasText && layout !== "imageOnly") {
    const textPoint = getTextPoint(position, imageBox, layout, gap, lineHeight, !hasImage);
    items.push({
      id: `${baseId}-text`,
      kind: "text",
      text,
      x: textPoint.x,
      y: textPoint.y,
      align: position.align,
      style,
    });
  }

  return items;
}

function getClampedImageBox(image: HeaderFooterZoneImage, page: PublishingPageLike, margin: number) {
  const maxWidth = Math.max(1, Math.min(image.maxWidth || image.width, page.width - margin * 2));
  const maxHeight = Math.max(1, Math.min(image.maxHeight || image.height, page.height / 5));
  let width = Math.max(1, Math.min(image.width, maxWidth));
  let height = Math.max(1, Math.min(image.height, maxHeight));
  if (image.maintainAspectRatio && image.width > 0 && image.height > 0) {
    const ratio = image.width / image.height;
    if (width / height > ratio) width = height * ratio;
    else height = width / ratio;
  }
  return { width, height };
}

function getImagePoint(position: { x: number; y: number; align: "left" | "center" | "right" }, imageBox: { width: number; height: number }, layout: string, gap: number, lineHeight: number, imageOnly: boolean, image: HeaderFooterZoneImage) {
  const verticalOffset = image.verticalAlign === "top" ? -imageBox.height / 2 : image.verticalAlign === "bottom" ? imageBox.height / 2 : 0;
  if (imageOnly) return { x: position.x, y: position.y + verticalOffset };
  if (layout === "textBeforeImage") return { x: position.x + horizontalSign(position.align) * (imageBox.width / 2 + gap + lineHeight), y: position.y + verticalOffset };
  if (layout === "imageBeforeText") return { x: position.x - horizontalSign(position.align) * (imageBox.width / 2 + gap + lineHeight), y: position.y + verticalOffset };
  if (layout === "textAboveImage") return { x: position.x, y: position.y + imageBox.height / 2 + gap + lineHeight / 2 + verticalOffset };
  if (layout === "imageAboveText") return { x: position.x, y: position.y - imageBox.height / 2 - gap - lineHeight / 2 + verticalOffset };
  return { x: position.x, y: position.y + verticalOffset };
}

function getTextPoint(position: { x: number; y: number; align: "left" | "center" | "right" }, imageBox: { width: number; height: number } | null, layout: string, gap: number, lineHeight: number, textOnly: boolean) {
  if (textOnly || !imageBox) return position;
  if (layout === "textBeforeImage") return { x: position.x - horizontalSign(position.align) * (imageBox.width / 2 + gap), y: position.y };
  if (layout === "imageBeforeText") return { x: position.x + horizontalSign(position.align) * (imageBox.width / 2 + gap), y: position.y };
  if (layout === "textAboveImage") return { x: position.x, y: position.y - imageBox.height / 2 - gap - lineHeight / 2 };
  if (layout === "imageAboveText") return { x: position.x, y: position.y + imageBox.height / 2 + gap + lineHeight / 2 };
  return position;
}

function horizontalSign(align: "left" | "center" | "right") {
  if (align === "right") return -1;
  return 1;
}

function unitless(value: number) {
  return Number.isFinite(value) ? value : 0;
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), Math.max(min, max));
}
