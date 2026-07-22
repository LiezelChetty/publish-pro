import { resolvePublishingPageIds } from "./pageRanges";
import { formatPageNumber, renderPublishingTokens } from "./tokens";
import type { PublishingPageLike, PublishingPreviewItem, PublishingSettings, PublishingTokenContext, PublishingUnit, PublishingZone } from "./types";

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
}: {
  settings: PublishingSettings;
  page: PublishingPageLike;
  pages: PublishingPageLike[];
  currentPage: number;
  selectedPageIds: string[];
  tokenContext: Omit<PublishingTokenContext, "pageNumber" | "pageCount" | "pageLabel">;
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
      const text = getHeaderFooterZoneText(settings, zone, page.pageNumber);
      if (!text.trim()) continue;
      const zonePosition = getZonePosition(zone, page, margin, 0);
      items.push({
        id: `${zone}-${page.id}`,
        kind: "text",
        text: renderPublishingTokens(text, context),
        ...zonePosition,
        style: settings.headerFooter.style,
      });
    }
  }

  if (settings.pageNumbers.enabled && resolvePublishingPageIds(settings.pageNumbers.target, pages, currentPage, selectedPageIds).has(page.id)) {
    const pageValue = settings.pageNumbers.startNumber + page.pageNumber - 1 + settings.pageNumbers.offset;
    const formatted = formatPageNumber(settings.pageNumbers.format, pageValue, pages.length, settings.pageNumbers.customTemplate);
    const template = settings.pageNumbers.format === "custom" ? formatted : `${settings.pageNumbers.prefix}${formatted}${settings.pageNumbers.suffix}`;
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
  const override = pageNumber === 1 ? settings.headerFooter.firstPage?.[zone] : pageNumber % 2 === 0 ? settings.headerFooter.evenPage?.[zone] : settings.headerFooter.oddPage?.[zone];
  if (settings.headerFooter.advanced && override !== undefined && override.trim() !== "") return override;
  const area = zone.startsWith("header") ? settings.headerFooter.header : settings.headerFooter.footer;
  const side = zone.endsWith("Left") ? "left" : zone.endsWith("Center") ? "center" : "right";
  return area[side].text;
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
