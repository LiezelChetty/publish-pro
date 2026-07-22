import type { DocumentBookmark } from "../navigation/types";
import type { PublishingSettings } from "../publishing/types";

export type DocxImportOptions = {
  fidelityMode: "fast" | "balanced" | "high";
  createBookmarks: boolean;
  preserveSource: boolean;
  importHeadersFooters: boolean;
  importHyperlinks: boolean;
  rebuildTocFromHeadings: boolean;
  fallbackPageSize: "a4" | "letter" | "legal";
  fallbackFont: "Helvetica" | "Times Roman" | "Courier";
  trackedChangesMode: "acceptAll" | "rejectAll" | "summary";
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
  id: string;
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  backgroundColor?: string;
  fontSize?: number;
  fontFamily?: string;
  letterSpacing?: number;
  verticalAlign?: "superscript" | "subscript" | "baseline";
  smallCaps?: boolean;
  allCaps?: boolean;
  hyperlink?: string;
};

export type DocxNoteReference = {
  id: string;
  kind: "footnote" | "endnote";
  marker: string;
};

export type DocxParagraphBlock = {
  id: string;
  type: "paragraph";
  sourcePath: string;
  sourceOrder: number;
  runs: DocxTextRun[];
  styleId?: string;
  styleName?: string;
  headingLevel?: number;
  alignment?: "left" | "center" | "right";
  indentLeft?: number;
  indentRight?: number;
  firstLineIndent?: number;
  hangingIndent?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: number;
  keepWithNext?: boolean;
  keepLinesTogether?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  list?: {
    numId: string;
    level: number;
    kind: "bullet" | "number" | "romanLower" | "romanUpper" | "alphaLower" | "alphaUpper";
    format?: string;
    text?: string;
    index: number;
  };
  noteReferences?: DocxNoteReference[];
  wordBookmarkNames?: string[];
  wordCommentIds?: string[];
};

export type DocxImageBlock = {
  id: string;
  type: "image";
  sourcePath: string;
  sourceOrder: number;
  relationshipId: string;
  altText?: string;
  width?: number;
  height?: number;
  floating?: boolean;
  wrapMode?: "inline" | "square" | "topBottom";
};

export type DocxTableCell = {
  blocks: DocxParagraphBlock[];
  background?: string;
  gridSpan?: number;
  width?: number;
  verticalAlign?: "top" | "middle" | "bottom";
};

export type DocxTableBlock = {
  id: string;
  type: "table";
  sourcePath: string;
  sourceOrder: number;
  rows: DocxTableCell[][];
  columnWidths?: number[];
  preferredWidth?: number;
  headerRows?: number[];
};

export type DocxPageBreakBlock = {
  id: string;
  type: "pageBreak";
  sourcePath: string;
  sourceOrder: number;
};

export type DocxBlock = DocxParagraphBlock | DocxImageBlock | DocxTableBlock | DocxPageBreakBlock;

export type DocxHeading = {
  blockId?: string;
  title: string;
  level: number;
  pageNumber: number;
};

export type DocxRenderedLink = {
  blockId: string;
  text: string;
  url: string;
  pageNumber: number;
  x: number;
  y: number;
  width: number;
  height: number;
  kind?: "external" | "internal";
  anchor?: string;
  destinationPageNumber?: number;
  destinationBlockId?: string;
  usedFallbackDestination?: boolean;
};

export type DocxSourceMapping = {
  blockId: string;
  kind: DocxBlock["type"];
  sourcePath: string;
  sourceOrder: number;
  pageNumber: number;
  x: number;
  y: number;
  width: number;
  height: number;
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
  category?: "typography" | "pagination" | "tables" | "images" | "hyperlinks" | "footnotes" | "comments" | "trackedChanges" | "sections" | "reimport" | "general";
  pageNumber?: number;
  sourceBlockId?: string;
};

export type DocxNote = {
  id: string;
  kind: "footnote" | "endnote";
  marker: string;
  text: string;
  hyperlinks: Array<{ text: string; url: string }>;
};

export type DocxWordBookmark = {
  name: string;
  blockId: string;
  sourceOrder: number;
};

export type DocxInternalLink = {
  blockId: string;
  anchor: string;
  text: string;
  sourceOrder: number;
};

export type DocxWordComment = {
  sourceId: string;
  author: string;
  initials?: string;
  date?: string;
  text: string;
  blockId?: string;
  sourceText?: string;
  approximate: boolean;
};

export type DocxTrackedChange = {
  id: string;
  type: "insertion" | "deletion" | "moveFrom" | "moveTo" | "formatting";
  author?: string;
  date?: string;
  text: string;
  blockId?: string;
};

export type DocxSectionNumbering = {
  blockId: string;
  sourceOrder: number;
  startValue: number;
  format: "decimal" | "romanLower" | "romanUpper" | "alphaLower" | "alphaUpper";
  prefix: string;
  unnumbered?: boolean;
};

export type DocxRevisionMetadata = {
  revisionId: string;
  sourceHash: string;
  importedAt: string;
  importerVersion: string;
  fidelityMode: DocxImportOptions["fidelityMode"];
  trackedChangesMode: DocxImportOptions["trackedChangesMode"];
  pageCount: number;
  warningCount: number;
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
  notes: DocxNote[];
  wordBookmarks: DocxWordBookmark[];
  internalLinks: DocxInternalLink[];
  wordComments: DocxWordComment[];
  trackedChanges: DocxTrackedChange[];
  sectionNumbering: DocxSectionNumbering[];
  statistics: {
    paragraphCount: number;
    tableCount: number;
    listCount: number;
    imageCount: number;
    footnoteCount: number;
    endnoteCount?: number;
    commentsDetected?: number;
    trackedChangesDetected?: number;
    sectionCount?: number;
    notesPlacedExactly?: number;
    notesUsingFallback?: number;
    brokenNoteReferences?: number;
    internalLinksDetected?: number;
    internalLinksMapped?: number;
    commentsImported?: number;
    approximateComments?: number;
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
  importId: string;
  fidelityMode: DocxImportOptions["fidelityMode"];
  pagesCreated: number;
  sectionsDetected: number;
  headingsFound: number;
  bookmarksCreated: number;
  imagesImported: number;
  tablesImported: number;
  listsImported: number;
  hyperlinksImported: number;
  headersFootersDetected: number;
  footnotesDetected: number;
  endnotesDetected: number;
  notesPlacedExactly: number;
  notesUsingFallback: number;
  brokenNoteReferences: number;
  commentsDetected: number;
  commentsImported: number;
  approximateComments: number;
  trackedChangesDetected: number;
  trackedChangesMode: DocxImportOptions["trackedChangesMode"];
  trackedChangeSummary: DocxTrackedChange[];
  internalLinksImported: number;
  internalLinkFallbacks: number;
  sectionNumberingDetected: number;
  revision: DocxRevisionMetadata;
  fontSubstitutions: string[];
  layoutSimplifications: string[];
  sourceMappings: number;
  parseTimeMs: number;
  renderTimeMs: number;
  totalTimeMs: number;
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
  links: DocxRenderedLink[];
  sourceMappings: DocxSourceMapping[];
  wordComments: DocxWordComment[];
  sectionNumbering: DocxSectionNumbering[];
  importMetadata: DocxImportMetadata;
  publishingSettingsPatch?: PublishingSettings;
  report: DocxImportReport;
};

export type DocxImportMetadata = {
  importId: string;
  sourceDocumentId?: string;
  originalSourceDocumentId?: string;
  sourceName: string;
  importedAt: string;
  fidelityMode: DocxImportOptions["fidelityMode"];
  options: DocxImportOptions;
  pageCount: number;
  warningCount: number;
  revisionHistory: DocxRevisionMetadata[];
};
