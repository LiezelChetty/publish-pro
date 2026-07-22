export type BookmarkStyle = {
  bold?: boolean;
  italic?: boolean;
  color?: string;
};

export type DocumentBookmark = {
  id: string;
  title: string;
  pageId: string;
  y?: number;
  parentId?: string;
  order: number;
  expanded: boolean;
  style?: BookmarkStyle;
  createdAt: string;
  updatedAt: string;
};

export type TocManualEntry = {
  id: string;
  title: string;
  pageId: string;
  level: number;
  order: number;
  enabled: boolean;
  createdAt: string;
  updatedAt: string;
};

export type TocInsertPosition = "beforeFirst" | "afterSelected";
export type TocSource = "bookmarks" | "pageLabels" | "manual";
export type TocPageSize = "a4" | "letter";
export type TocFontFamily = "Helvetica" | "Times Roman" | "Courier";

export type TocSettings = {
  title: string;
  source: TocSource;
  includePageLabels: boolean;
  includePageNumbers: boolean;
  dotLeaders: boolean;
  maxDepth: number;
  indent: number;
  lineSpacing: number;
  fontFamily: TocFontFamily;
  fontSize: number;
  color: string;
  boldLevels: number[];
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  pageSize: TocPageSize;
  insertPosition: TocInsertPosition;
  includeTocInNumbering: boolean;
  startMainNumberingAfterToc: boolean;
};

export type TocGeneratedPageMetadata = {
  pageIds: string[];
  insertedAtPageId?: string;
  generatedAt: string;
  stale: boolean;
};

export type NavigationState = {
  bookmarks: DocumentBookmark[];
  tocSettings: TocSettings;
  manualTocEntries: TocManualEntry[];
  generatedToc?: TocGeneratedPageMetadata;
};

export type PageReference = {
  id: string;
  pageNumber: number;
  label?: string;
  width: number;
  height: number;
};

export type TocEntry = {
  id: string;
  title: string;
  pageId: string;
  level: number;
  source: "bookmark" | "pageLabel" | "manual";
};

export type TocLine = TocEntry & {
  pageText: string;
  text: string;
  x: number;
  y: number;
  pageIndex: number;
  width: number;
  height: number;
};

export type TocLayoutPage = {
  id: string;
  lines: TocLine[];
};

export type TocLayoutResult = {
  pages: TocLayoutPage[];
  width: number;
  height: number;
};
