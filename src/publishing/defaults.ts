import type { PublishingSettings, PublishingTextStyle } from "./types";

export const publishingWatermarkPresets = ["DRAFT", "CONFIDENTIAL", "FOR REVIEW", "APPROVED", "NOT FOR CONSTRUCTION"];

export const defaultPublishingTextStyle: PublishingTextStyle = {
  fontFamily: "Helvetica",
  fontSize: 10,
  bold: false,
  italic: false,
  color: "#111827",
  opacity: 0.82,
};

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
        left: { text: "{project}" },
        center: { text: "" },
        right: { text: "{date}" },
      },
      footer: {
        left: { text: "{client}" },
        center: { text: "" },
        right: { text: "{filename}" },
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
        left: { ...defaults.headerFooter.header.left, ...input.headerFooter?.header?.left },
        center: { ...defaults.headerFooter.header.center, ...input.headerFooter?.header?.center },
        right: { ...defaults.headerFooter.header.right, ...input.headerFooter?.header?.right },
      },
      footer: {
        left: { ...defaults.headerFooter.footer.left, ...input.headerFooter?.footer?.left },
        center: { ...defaults.headerFooter.footer.center, ...input.headerFooter?.footer?.center },
        right: { ...defaults.headerFooter.footer.right, ...input.headerFooter?.footer?.right },
      },
    },
    watermark: {
      ...defaults.watermark,
      ...input.watermark,
      target: { ...defaults.watermark.target, ...input.watermark?.target },
      style: { ...defaults.watermark.style, ...input.watermark?.style },
    },
  };
}
