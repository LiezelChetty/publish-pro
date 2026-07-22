export type PublishingZone = "headerLeft" | "headerCenter" | "headerRight" | "footerLeft" | "footerCenter" | "footerRight";
export type PublishingTargetMode = "all" | "selected" | "current" | "custom" | "odd" | "even" | "exceptFirst" | "exceptSelected";
export type PublishingNumberFormat = "decimal" | "decimal2" | "decimal3" | "romanLower" | "romanUpper" | "alphaLower" | "alphaUpper" | "page" | "pageOfPages" | "custom";
export type PublishingFontFamily = "Helvetica" | "Times Roman" | "Courier";
export type PublishingUnit = "mm" | "cm" | "in" | "pt";
export type PublishingPositionPreset = "center" | "topLeft" | "topCenter" | "topRight" | "bottomLeft" | "bottomCenter" | "bottomRight" | "custom";
export type PublishingZoneImageLayout = "imageOnly" | "textBeforeImage" | "imageBeforeText" | "imageAboveText" | "textAboveImage";
export type PublishingHorizontalAlign = "left" | "center" | "right";
export type PublishingVerticalAlign = "top" | "middle" | "bottom";

export type PublishingTarget = {
  mode: PublishingTargetMode;
  range: string;
  excludeFirst: boolean;
};

export type PublishingTextStyle = {
  fontFamily: PublishingFontFamily;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  color: string;
  opacity: number;
};

export type HeaderFooterZone = {
  text: string;
  image?: HeaderFooterZoneImage;
};

export type HeaderFooterZoneImage = {
  assetId?: string;
  width: number;
  height: number;
  maxWidth: number;
  maxHeight: number;
  maintainAspectRatio: boolean;
  opacity: number;
  horizontalAlign: PublishingHorizontalAlign;
  verticalAlign: PublishingVerticalAlign;
  padding: number;
  offsetX: number;
  offsetY: number;
  layout: PublishingZoneImageLayout;
};

export type HeaderFooterZoneOverride = string | HeaderFooterZone;

export type HeaderFooterSettings = {
  enabled: boolean;
  advanced: boolean;
  target: PublishingTarget;
  unit: PublishingUnit;
  margin: number;
  style: PublishingTextStyle;
  header: Record<"left" | "center" | "right", HeaderFooterZone>;
  footer: Record<"left" | "center" | "right", HeaderFooterZone>;
  firstPage?: Partial<Record<PublishingZone, HeaderFooterZoneOverride>>;
  oddPage?: Partial<Record<PublishingZone, HeaderFooterZoneOverride>>;
  evenPage?: Partial<Record<PublishingZone, HeaderFooterZoneOverride>>;
};

export type PageNumberSettings = {
  enabled: boolean;
  zone: PublishingZone;
  format: PublishingNumberFormat;
  customTemplate: string;
  startNumber: number;
  offset: number;
  prefix: string;
  suffix: string;
  target: PublishingTarget;
  unit: PublishingUnit;
  distanceFromEdge: number;
  horizontalOffset: number;
  style: PublishingTextStyle;
};

export type WatermarkSettings = {
  enabled: boolean;
  type: "text" | "image";
  text: string;
  imageAssetId?: string;
  preset: string;
  target: PublishingTarget;
  position: PublishingPositionPreset;
  customX: number;
  customY: number;
  unit: PublishingUnit;
  opacity: number;
  rotation: number;
  scale: number;
  maintainAspectRatio: boolean;
  tiled: boolean;
  style: PublishingTextStyle;
};

export type PublishingSettings = {
  pageNumbers: PageNumberSettings;
  headerFooter: HeaderFooterSettings;
  watermark: WatermarkSettings;
};

export type PublishingTokenContext = {
  pageNumber: number;
  pageCount: number;
  pageLabel?: string;
  projectName: string;
  client: string;
  filename: string;
  date: string;
};

export type PublishingPageLike = {
  id: string;
  pageNumber: number;
  width: number;
  height: number;
  label?: string;
};

export type PublishingPreviewItem = {
  id: string;
  kind: "text" | "watermark" | "safeArea" | "image" | "missingImage";
  text?: string;
  assetId?: string;
  x: number;
  y: number;
  width?: number;
  height?: number;
  align?: "left" | "center" | "right";
  rotation?: number;
  style?: PublishingTextStyle;
  opacity?: number;
};
