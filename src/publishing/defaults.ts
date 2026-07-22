import type { HeaderFooterZone, HeaderFooterZoneImage, HeaderFooterZoneOverride, PublishingSettings, PublishingTextStyle } from "./types";

export const publishingWatermarkPresets = ["DRAFT", "CONFIDENTIAL", "FOR REVIEW", "APPROVED", "NOT FOR CONSTRUCTION"];

export const defaultPublishingTextStyle: PublishingTextStyle = {
  fontFamily: "Helvetica",
  fontSize: 10,
  bold: false,
  italic: false,
  color: "#111827",
  opacity: 0.82,
};

export function createDefaultHeaderFooterImage(): HeaderFooterZoneImage {
  return {
    assetId: undefined,
    width: 48,
    height: 24,
    maxWidth: 120,
    maxHeight: 48,
    maintainAspectRatio: true,
    opacity: 1,
    horizontalAlign: "center",
    verticalAlign: "middle",
    padding: 6,
    offsetX: 0,
    offsetY: 0,
    layout: "imageBeforeText",
  };
}

export function mergeHeaderFooterZone(value: unknown, fallbackText = ""): HeaderFooterZone {
  const defaults: HeaderFooterZone = { text: fallbackText };
  if (typeof value === "string") return { text: value };
  if (!value || typeof value !== "object") return defaults;
  const input = value as Partial<HeaderFooterZone>;
  return {
    text: input.text ?? fallbackText,
    image: input.image ? { ...createDefaultHeaderFooterImage(), ...input.image } : undefined,
  };
}

export function mergeHeaderFooterOverride(value: unknown): HeaderFooterZoneOverride | undefined {
  if (typeof value === "string") return value;
  if (!value || typeof value !== "object") return undefined;
  return mergeHeaderFooterZone(value);
}

export function createDefaultPublishingSettings(): PublishingSettings {
  return {
    pageNumbers: {
      enabled: false,
      zone: "footerCenter",
      format: "pageOfPages",
      customTemplate: "Page {page} of {pages}",
      startNumber: 1,
      offset: 0,
      prefix: "",
      suffix: "",
      target: { mode: "all", range: "", excludeFirst: false },
      unit: "mm",
      distanceFromEdge: 12,
      horizontalOffset: 0,
      style: { ...defaultPublishingTextStyle },
    },
    headerFooter: {
      enabled: false,
      advanced: false,
      target: { mode: "all", range: "", excludeFirst: false },
      unit: "mm",
      margin: 12,
      style: { ...defaultPublishingTextStyle },
      header: {
        left: mergeHeaderFooterZone(undefined, "{project}"),
        center: mergeHeaderFooterZone(undefined, ""),
        right: mergeHeaderFooterZone(undefined, "{date}"),
      },
      footer: {
        left: mergeHeaderFooterZone(undefined, "{client}"),
        center: mergeHeaderFooterZone(undefined, ""),
        right: mergeHeaderFooterZone(undefined, "{filename}"),
      },
    },
    watermark: {
      enabled: false,
      type: "text",
      text: "DRAFT",
      preset: "DRAFT",
      target: { mode: "all", range: "", excludeFirst: false },
      position: "center",
      customX: 0,
      customY: 0,
      unit: "mm",
      opacity: 0.16,
      rotation: -35,
      scale: 1,
      maintainAspectRatio: true,
      tiled: false,
      style: {
        ...defaultPublishingTextStyle,
        fontSize: 64,
        bold: true,
        color: "#6b7280",
      },
    },
  };
}

export function mergePublishingSettings(value: unknown): PublishingSettings {
  const defaults = createDefaultPublishingSettings();
  if (!value || typeof value !== "object") return defaults;
  const input = value as Partial<PublishingSettings>;
  return {
    pageNumbers: {
      ...defaults.pageNumbers,
      ...input.pageNumbers,
      target: { ...defaults.pageNumbers.target, ...input.pageNumbers?.target },
      style: { ...defaults.pageNumbers.style, ...input.pageNumbers?.style },
    },
    headerFooter: {
      ...defaults.headerFooter,
      ...input.headerFooter,
      target: { ...defaults.headerFooter.target, ...input.headerFooter?.target },
      style: { ...defaults.headerFooter.style, ...input.headerFooter?.style },
      header: {
        left: mergeHeaderFooterZone(input.headerFooter?.header?.left, defaults.headerFooter.header.left.text),
        center: mergeHeaderFooterZone(input.headerFooter?.header?.center, defaults.headerFooter.header.center.text),
        right: mergeHeaderFooterZone(input.headerFooter?.header?.right, defaults.headerFooter.header.right.text),
      },
      footer: {
        left: mergeHeaderFooterZone(input.headerFooter?.footer?.left, defaults.headerFooter.footer.left.text),
        center: mergeHeaderFooterZone(input.headerFooter?.footer?.center, defaults.headerFooter.footer.center.text),
        right: mergeHeaderFooterZone(input.headerFooter?.footer?.right, defaults.headerFooter.footer.right.text),
      },
      firstPage: mergeHeaderFooterOverrideMap(input.headerFooter?.firstPage),
      oddPage: mergeHeaderFooterOverrideMap(input.headerFooter?.oddPage),
      evenPage: mergeHeaderFooterOverrideMap(input.headerFooter?.evenPage),
    },
    watermark: {
      ...defaults.watermark,
      ...input.watermark,
      target: { ...defaults.watermark.target, ...input.watermark?.target },
      style: { ...defaults.watermark.style, ...input.watermark?.style },
    },
  };
}

function mergeHeaderFooterOverrideMap(value: unknown) {
  if (!value || typeof value !== "object") return undefined;
  const input = value as Record<string, unknown>;
  const output: Record<string, HeaderFooterZoneOverride> = {};
  for (const [zone, override] of Object.entries(input)) {
    const merged = mergeHeaderFooterOverride(override);
    if (merged !== undefined) output[zone] = merged;
  }
  return Object.keys(output).length > 0 ? output : undefined;
}
