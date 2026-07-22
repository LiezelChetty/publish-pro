import type { DocumentBookmark } from "../navigation/types";
import type { PublishingSettings } from "../publishing/types";

export type DocxImportOptions = {
  createBookmarks: boolean;
  preserveSource: boolean;
  importHeadersFooters: boolean;
  importHyperlinks: boolean;
  rebuildTocFromHeadings: boolean;
  fallbackPageSize: "a4" | "letter" | "legal";
  fallbackFont: "Helvetica" | "Times Roman" | "Courier";
};

export type DocxImportProgress = {
  stage: string;
  progress: number;
};

export type DocxPageSize = {
  width: number;
  height: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
};

export type DocxTextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  fontSize?: number;
  hyperlink?: string;
};

export type DocxParagraphBlock = {
  type: "paragraph";
  runs: DocxTextRun[];
  styleId?: string;
  headingLevel?: number;
  alignment?: "left" | "center" | "right";
  indentLeft?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  list?: {
    level: number;
    kind: "bullet" | "number";
    index: number;
  };
};

export type DocxImageBlock = {
  type: "image";
  relationshipId: string;
  altText?: string;
  width?: number;
  height?: number;
};

export type DocxTableCell = {
  blocks: DocxParagraphBlock[];
  background?: string;
};

export type DocxTableBlock = {
  type: "table";
  rows: DocxTableCell[][];
};

export type DocxPageBreakBlock = {
  type: "pageBreak";
};

export type DocxBlock = DocxParagraphBlock | DocxImageBlock | DocxTableBlock | DocxPageBreakBlock;

export type DocxHeading = {
  title: string;
  level: number;
  pageNumber: number;
};

export type DocxExtractedImage = {
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

export type DocxImportWarning = {
  code: string;
  message: string;
};

export type DocxIntermediateDocument = {
  title: string;
  blocks: DocxBlock[];
  pageSize: DocxPageSize;
  images: DocxExtractedImage[];
  headings: DocxHeading[];
  headerText?: string;
  footerText?: string;
  hyperlinks: Array<{ text: string; url: string }>;
  statistics: {
    paragraphCount: number;
    tableCount: number;
    listCount: number;
    imageCount: number;
    footnoteCount: number;
  };
  warnings: DocxImportWarning[];
};

export type DocxImportedSource = {
  id: string;
  name: string;
  bytes: Uint8Array;
  mimeType: string;
};

export type DocxImportReport = {
  pagesCreated: number;
  headingsFound: number;
  bookmarksCreated: number;
  imagesImported: number;
  tablesImported: number;
  listsImported: number;
  headersFootersDetected: number;
  footnotesDetected: number;
  warnings: DocxImportWarning[];
  unsupportedContent: string[];
};

export type DocxImportResult = {
  projectName: string;
  defaultPdfName: string;
  convertedPdfSource: DocxImportedSource;
  originalSource?: DocxImportedSource;
  imageAssets: DocxExtractedImage[];
  headings: DocxHeading[];
  bookmarks: DocumentBookmark[];
  publishingSettingsPatch?: PublishingSettings;
  report: DocxImportReport;
};
