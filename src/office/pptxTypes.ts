import type { DocumentBookmark } from "../navigation/types";
import type { DocxImportedSource, DocxImportProgress, DocxImportWarning } from "./types";

export type PptxImportOptions = {
  fidelityMode: "fast" | "balanced" | "high";
  preserveSource: boolean;
  createBookmarks: boolean;
  includeHiddenSlides: boolean;
};

export type PptxImportProgress = DocxImportProgress;

export type PptxRelationship = {
  type: string;
  target: string;
  external: boolean;
};

export type PptxPackage = {
  files: Record<string, Uint8Array>;
  textFiles: Map<string, string>;
  warnings: DocxImportWarning[];
};

export type PptxSlideSize = {
  width: number;
  height: number;
};

export type PptxTextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  fontSize?: number;
  fontFamily?: string;
};

export type PptxTextBlock = {
  id: string;
  kind: "text";
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  paragraphs: PptxTextRun[][];
  alignment?: "left" | "center" | "right";
  sourcePath: string;
};

export type PptxShapeElement = {
  id: string;
  kind: "shape";
  preset: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  sourcePath: string;
};

export type PptxImageElement = {
  id: string;
  kind: "image";
  relationshipId: string;
  assetId?: string;
  altText?: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  sourcePath: string;
};

export type PptxTableElement = {
  id: string;
  kind: "table";
  x: number;
  y: number;
  width: number;
  height: number;
  rows: string[][];
  sourcePath: string;
};

export type PptxChartElement = {
  id: string;
  kind: "chart";
  x: number;
  y: number;
  width: number;
  height: number;
  chartType?: string;
  sourcePath: string;
};

export type PptxSlideElement = PptxTextBlock | PptxShapeElement | PptxImageElement | PptxTableElement | PptxChartElement;

export type PptxImageAsset = {
  id: string;
  relationshipId: string;
  name: string;
  path: string;
  mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/svg+xml" | "image/unsupported";
  bytes: Uint8Array;
  dataUrl: string;
  width: number;
  height: number;
  used: boolean;
};

export type PptxHyperlink = {
  slideIndex: number;
  elementId: string;
  text: string;
  url?: string;
  targetSlideIndex?: number;
  x: number;
  y: number;
  width: number;
  height: number;
};

export type PptxSlide = {
  id: string;
  sourceSlideId: string;
  slideIndex: number;
  title: string;
  hidden: boolean;
  sourcePath: string;
  layoutTarget?: string;
  masterTarget?: string;
  background?: string;
  elements: PptxSlideElement[];
  notes: string;
  hyperlinks: PptxHyperlink[];
};

export type PptxSection = {
  id: string;
  name: string;
  firstSlideId?: string;
};

export type PptxIntermediatePresentation = {
  title: string;
  slideSize: PptxSlideSize;
  slides: PptxSlide[];
  sections: PptxSection[];
  imageAssets: PptxImageAsset[];
  warnings: DocxImportWarning[];
  metadata: {
    created?: string;
    modified?: string;
    creator?: string;
    slideCount: number;
  };
  statistics: {
    hiddenSlides: number;
    textBlocks: number;
    shapes: number;
    images: number;
    tables: number;
    charts: number;
    speakerNotes: number;
    hyperlinks: number;
    rasterisedElements: number;
    unsupportedElements: number;
  };
};

export type PptxSourceMapping = {
  slideId: string;
  elementId: string;
  kind: PptxSlideElement["kind"] | "slide";
  sourcePath: string;
  pageNumber: number;
  x: number;
  y: number;
  width: number;
  height: number;
};

export type PptxRevisionMetadata = {
  revisionId: string;
  sourceHash: string;
  importedAt: string;
  importerVersion: string;
  fidelityMode: PptxImportOptions["fidelityMode"];
  slideCount: number;
  warningCount: number;
};

export type PptxImportMetadata = {
  kind: "pptx";
  importId: string;
  sourceDocumentId?: string;
  originalSourceDocumentId?: string;
  sourceName: string;
  importedAt: string;
  fidelityMode: PptxImportOptions["fidelityMode"];
  options: PptxImportOptions;
  pageCount: number;
  warningCount: number;
  revisionHistory: PptxRevisionMetadata[];
  slideTitles: Array<{ pageNumber: number; title: string; hidden: boolean }>;
  speakerNotes: Array<{ pageNumber: number; slideId: string; title: string; notes: string }>;
  sections: PptxSection[];
};

export type PptxImportReport = {
  importId: string;
  fidelityMode: PptxImportOptions["fidelityMode"];
  slidesImported: number;
  hiddenSlides: number;
  slideSections: number;
  textBoxes: number;
  images: number;
  shapes: number;
  tables: number;
  charts: number;
  smartArtFallbacks: number;
  speakerNotes: number;
  hyperlinks: number;
  fontSubstitutions: string[];
  rasterisedElements: number;
  unsupportedElements: number;
  parseTimeMs: number;
  renderTimeMs: number;
  totalTimeMs: number;
  revision: PptxRevisionMetadata;
  warnings: DocxImportWarning[];
};

export type PptxImportResult = {
  projectName: string;
  defaultPdfName: string;
  convertedPdfSource: DocxImportedSource;
  originalSource?: DocxImportedSource;
  imageAssets: PptxImageAsset[];
  bookmarks: DocumentBookmark[];
  sourceMappings: PptxSourceMapping[];
  importMetadata: PptxImportMetadata;
  report: PptxImportReport;
};
