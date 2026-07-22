import {
  BookOpen,
  CheckCircle2,
  AlertTriangle,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  ChevronUp,
  Download,
  FilePlus2,
  Files,
  Hand,
  Highlighter,
  Image as ImageIcon,
  Circle,
  CornerUpRight,
  Maximize2,
  MessageSquare,
  Minus,
  MousePointer2,
  MoreHorizontal,
  PanelLeftClose,
  PanelLeftOpen,
  PanelRightClose,
  PanelRightOpen,
  PenLine,
  Pencil,
  Plus,
  Redo2,
  Search,
  Square,
  Stamp,
  Strikethrough,
  TextCursorInput,
  Trash2,
  Undo2,
  Underline,
  ZoomIn,
  ZoomOut,
} from "lucide-react";
import { degrees, PDFDocument, rgb, StandardFonts, type PDFFont, type PDFPage } from "pdf-lib";
import {
  GlobalWorkerOptions,
  getDocument,
  type PDFDocumentProxy,
  type PDFPageProxy,
} from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";
import { ChangeEvent, PointerEvent, ReactElement, type DragEvent as ReactDragEvent, type KeyboardEvent as ReactKeyboardEvent, type MouseEvent as ReactMouseEvent, useEffect, useMemo, useRef, useState } from "react";
import {
  commentMatchesFilter,
  createCommentData,
  createCommentReply,
  formatCommentDate,
  getCommentPreview,
  wrapText,
  type CommentData,
  type CommentReply,
  type CommentStatusFilter,
} from "./comments";
import {
  getMarkupBounds,
  isTextMarkupKind,
  moveMarkupRects,
  normalizeSelectionRects,
  scaleMarkupRect,
  type MarkupRect,
  type TextMarkupKind,
} from "./textMarkup";
import {
  createDefaultTextBoxStyle,
  getCssFontFamily,
  getPresetTextStyle,
  getStandardFontName,
  measureTextWidth,
  textBoxAlignments,
  textBoxFontFamilies,
  textBoxPresets,
  wrapTextForBox,
  type TextBoxFontFamily,
  type TextBoxPreset,
  type TextBoxStyle,
} from "./textBoxes";
import {
  createShapeStyle,
  getDashArray,
  getPdfDashArray,
  shapeDashStyles,
  shapeToolLabels,
  type ShapeKind,
  type ShapeStyle,
} from "./shapes";
import { AppChrome, type ThemeMode } from "./components/app-shell/AppChrome";
import { createDefaultHeaderFooterImage, createDefaultPublishingSettings, mergeHeaderFooterZone, mergePublishingSettings, publishingWatermarkPresets } from "./publishing/defaults";
import { drawPublishingMarksToPdf } from "./publishing/export";
import { getPublishingPreviewItems } from "./publishing/layout";
import { createNumberingSection, duplicateNumberingSection, type NumberingSection } from "./publishing/numberingSections";
import { defaultLegacyNumberingSection, getNumberingSummary, getResolvedPageText, resolveNumberingSections } from "./publishing/numberingResolver";
import { validateNumberingSections } from "./publishing/numberingValidation";
import { getPageRangeError } from "./publishing/pageRanges";
import { deletePublishingPreset, loadPublishingPresets, savePublishingPreset, type PublishingPreset } from "./publishing/presets";
import type { HeaderFooterZone, HeaderFooterZoneImage, PublishingNumberFormat, PublishingPositionPreset, PublishingSettings, PublishingTargetMode, PublishingUnit, PublishingZone, PublishingZoneImageLayout } from "./publishing/types";
import { collectProjectAssets, estimateDataUrlSize, formatAssetSize, type ProjectImageAssetLike } from "./projects/assets";
import { createProjectMetadata, type ProjectMetadata } from "./projects/schema";
import { buildProjectManifest, deserializeProject, serializeProject } from "./projects/serialization";
import {
  createBookmark,
  createDefaultTocSettings,
  deleteBookmark,
  duplicateBookmarkTree,
  getBookmarkDescendants,
  getBookmarkLevel,
  getVisibleBookmarkTree,
  moveBookmark,
  nestBookmark,
  normalizeBookmarks,
  outdentBookmark,
  wouldCreateBookmarkCycle,
} from "./navigation/bookmarks";
import { addPdfOutline, addTocLinks } from "./navigation/pdfOutline";
import { buildTocEntries, getTocPageSize, layoutToc } from "./navigation/toc";
import { validateNavigation } from "./navigation/validation";
import { validateExportLinks } from "./navigation/linkValidation";
import type { DocumentBookmark, NavigationState, PageReference, TocGeneratedPageMetadata, TocLayoutResult, TocManualEntry, TocSettings } from "./navigation/types";
import { createBookmarksFromDocxHeadings } from "./office/docxBookmarks";
import { importDocxDocument } from "./office/docxImport";
import { defaultDocxImportOptions, validateDocxImportFile } from "./office/validation";
import type { DocxImportMetadata, DocxImportOptions, DocxImportReport, DocxSectionNumbering, DocxSourceMapping, DocxWordComment } from "./office/types";
import { importPptxPresentation } from "./office/pptxImport";
import { defaultPptxImportOptions, validatePptxImportFile } from "./office/pptxValidation";
import type { PptxImportMetadata, PptxImportOptions, PptxImportReport, PptxSourceMapping } from "./office/pptxTypes";
import { runtime, runtimeFileToBrowserFile } from "./runtime";
import {
  addRecentProject as addBrowserRecentProject,
  clearRecentProjects as clearBrowserRecentProjects,
  clearAutosave as clearBrowserAutosave,
  inputToMetadataTags,
  loadAutosave as loadBrowserAutosave,
  loadRecentProjects as loadBrowserRecentProjects,
  metadataTagsToInput,
  removeRecentProject as removeBrowserRecentProject,
  saveAutosave as saveBrowserAutosave,
  type ProjectAutosave,
  type RecentProject,
} from "./projects/storage";

GlobalWorkerOptions.workerSrc = pdfWorker;

const THEME_STORAGE_KEY = "publish-pro-theme";
const WORKSPACE_STORAGE_KEY = "publish-pro-active-workspace";

type ShapeTool = `shape:${ShapeKind}`;
type Tool = "select" | "text" | "highlight" | "signature" | "stamp" | "image" | "draw" | "pngSignature" | "comment" | ShapeTool | TextMarkupKind;
type MarkKind = "text" | "highlight" | "signature" | "stamp" | "image" | "stroke" | "pngSignature" | "comment" | "shape" | TextMarkupKind;
type LeftPanel = "pages" | "insert" | "bookmarks" | "comments" | "search";
type WorkspaceMode = "import" | "assemble" | "review" | "publish";
type WorkState = {
  message: string;
  progress?: number;
};
type ToastTone = "success" | "info" | "warning" | "error" | "progress";
type ToastMessage = {
  id: string;
  tone: ToastTone;
  title: string;
  detail?: string;
};

type StrokePoint = {
  x: number;
  y: number;
};

type PreparedImage = {
  dataUrl: string;
  mimeType: "image/png" | "image/jpeg";
  name: string;
  width: number;
  height: number;
};

type ProjectImageAsset = ProjectImageAssetLike & {
  mimeType: "image/png" | "image/jpeg";
  naturalWidth: number;
  naturalHeight: number;
  updatedAt: string;
};

type OfficeImportMetadata = DocxImportMetadata | PptxImportMetadata;
type OfficeSourceMapping = DocxSourceMapping | PptxSourceMapping;

type ProjectAssetImportTarget =
  | { kind: "projectAssets" }
  | { kind: "watermark" }
  | { kind: "headerFooter"; area: "header" | "footer"; side: "left" | "center" | "right" }
  | { kind: "headerFooterOverride"; key: "firstPage" | "oddPage" | "evenPage"; zone: PublishingZone };

type TextItemView = {
  id: string;
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  fontSize: number;
};

type Mark = {
  id: string;
  kind: MarkKind;
  pageId?: string;
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  color: string;
  size: number;
  rotation: number;
  imageDataUrl?: string;
  imageMimeType?: "image/png" | "image/jpeg";
  imageName?: string;
  imageNaturalWidth?: number;
  imageNaturalHeight?: number;
  opacity?: number;
  thickness?: number;
  lockAspectRatio?: boolean;
  markupRects?: MarkupRect[];
  comment?: CommentData;
  textStyle?: TextBoxStyle;
  shapeStyle?: ShapeStyle;
  strokePoints?: StrokePoint[];
  strokeOpacity?: number;
};

type HistoryEntry = {
  pagesBefore: PageView[];
  pagesAfter: PageView[];
  marksBefore: Mark[];
  marksAfter: Mark[];
  selectedBefore: string | null;
  selectedAfter: string | null;
  currentPageBefore: number;
  currentPageAfter: number;
  selectedPageIdsBefore: string[];
  selectedPageIdsAfter: string[];
};

type HistoryState = {
  past: HistoryEntry[];
  future: HistoryEntry[];
};

type NavigationHistoryState = {
  past: NavigationState[];
  future: NavigationState[];
};

type PageView = {
  id: string;
  pageNumber: number;
  sourceDocumentId?: string;
  sourcePageNumber?: number;
  importMetadata?: OfficeImportMetadata;
  sourceMappings?: OfficeSourceMapping[];
  width: number;
  height: number;
  rotation: number;
  label?: string;
  isBlank?: boolean;
  generatedToc?: boolean;
  background?: string;
  imageUrl: string;
  textItems: TextItemView[];
};

type SourceDocument = {
  id: string;
  name: string;
  bytes: Uint8Array;
  mimeType?: string;
  importMetadata?: OfficeImportMetadata;
};

type BlankPagePreset = "a4Portrait" | "a4Landscape" | "a3Portrait" | "a3Landscape" | "letterPortrait" | "letterLandscape" | "legalPortrait" | "legalLandscape" | "matchCurrent" | "custom";
type PageDialog = "blank" | "merge" | "replace" | "split" | "extract" | "labels" | null;
type InsertPosition = "start" | "before" | "after" | "end";
type DimensionUnit = "mm" | "cm" | "in" | "pt";
type PageOrientation = "portrait" | "landscape";
type SplitMode = "afterCurrent" | "everyN" | "ranges" | "selected" | "selectedBoundaries";

type MergeQueueItem = {
  id: string;
  file: File;
  name: string;
  bytes: Uint8Array;
  sourceDocumentId: string;
  pdfDoc: PDFDocumentProxy;
  pageCount: number;
  range: string;
  error?: string;
};

type DocxReimportImpact = {
  currentPages: number;
  currentBookmarks: number;
  manualAnnotationsAffected: number;
  commentsAffected: number;
  manualBookmarksPreserved: number;
  publishingSettingsPreserved: boolean;
};

type PptxReimportImpact = {
  currentSlides: number;
  currentBookmarks: number;
  manualAnnotationsAffected: number;
  speakerNotes: number;
  publishingSettingsPreserved: boolean;
};

type PageContextMenuState = {
  x: number;
  y: number;
  pageId: string;
} | null;

const PAGE_SCALE = 1.35;
const MAX_HISTORY_ENTRIES = 100;
const BRAND_RED = "#d8342a";
const BRAND_ICON_SRC = "/brand/publish-pro-icon.svg";
const COMMENT_MARKER_SIZE = 32;

const demoPages: PageView[] = [
  {
    id: "demo-page-1",
    pageNumber: 1,
    rotation: 0,
    isBlank: true,
    width: 612,
    height: 792,
    imageUrl: "",
    textItems: [],
  },
];

const initialMarks: Mark[] = [
  {
    id: "demo-title",
    kind: "text",
    page: 1,
    x: 72,
    y: 92,
    width: 300,
    height: 34,
    text: "Drop a PDF to begin",
    color: "#1f2937",
    size: 24,
    rotation: 0,
  },
  {
    id: "demo-note",
    kind: "stamp",
    page: 1,
    x: 72,
    y: 150,
    width: 132,
    height: 34,
    text: "READY",
    color: BRAND_RED,
    size: 16,
    rotation: 0,
  },
];

export function App() {
  const [themeMode, setThemeMode] = useState<ThemeMode>(() => getStoredThemeMode());
  const [resolvedTheme, setResolvedTheme] = useState<"light" | "dark">(() => resolveThemeMode(getStoredThemeMode()));
  const [hasOpenProject, setHasOpenProject] = useState(false);
  const [projectId, setProjectId] = useState<string>(() => crypto.randomUUID());
  const [projectMetadata, setProjectMetadata] = useState<ProjectMetadata>(() => createProjectMetadata());
  const [isProjectDirty, setIsProjectDirty] = useState(false);
  const [savedProjectFingerprint, setSavedProjectFingerprint] = useState("");
  const [recentProjects, setRecentProjects] = useState<RecentProject[]>(() => runtime.isDesktop ? [] : loadBrowserRecentProjects());
  const [autosaveCandidate, setAutosaveCandidate] = useState<ProjectAutosave | null>(null);
  const [desktopProjectPath, setDesktopProjectPath] = useState<string | null>(null);
  const [lastExportPath, setLastExportPath] = useState<string | null>(null);
  const [publishingSettings, setPublishingSettings] = useState<PublishingSettings>(() => createDefaultPublishingSettings());
  const [publishingHistory, setPublishingHistory] = useState<{ past: PublishingSettings[]; future: PublishingSettings[] }>({ past: [], future: [] });
  const [publishingPresets, setPublishingPresets] = useState<PublishingPreset[]>(() => loadPublishingPresets());
  const [publishingValidationMessage, setPublishingValidationMessage] = useState("");
  const [lastSavedAt, setLastSavedAt] = useState<string | null>(null);
  const [lastAutosavedAt, setLastAutosavedAt] = useState<string | null>(null);
  const [projectStatusMessage, setProjectStatusMessage] = useState("");
  const [isProjectDialogOpen, setIsProjectDialogOpen] = useState(false);
  const [newProjectDraft, setNewProjectDraft] = useState(() => createProjectMetadata());
  const [pdfName, setPdfName] = useState("Untitled document");
  const [pdfBytes, setPdfBytes] = useState<Uint8Array | null>(null);
  const [pdfDoc, setPdfDoc] = useState<PDFDocumentProxy | null>(null);
  const [activeSourceDocumentId, setActiveSourceDocumentId] = useState<string | null>(null);
  const [sourceDocuments, setSourceDocuments] = useState<Record<string, SourceDocument>>({});
  const [projectImageAssets, setProjectImageAssets] = useState<ProjectImageAsset[]>([]);
  const [pages, setPages] = useState<PageView[]>(demoPages);
  const [currentPage, setCurrentPage] = useState(1);
  const [marks, setMarks] = useState<Mark[]>(initialMarks);
  const [history, setHistory] = useState<HistoryState>({ past: [], future: [] });
  const [bookmarks, setBookmarks] = useState<DocumentBookmark[]>([]);
  const [selectedBookmarkId, setSelectedBookmarkId] = useState<string | null>(null);
  const [tocSettings, setTocSettings] = useState<TocSettings>(createDefaultTocSettings());
  const [manualTocEntries, setManualTocEntries] = useState<TocManualEntry[]>([]);
  const [generatedToc, setGeneratedToc] = useState<TocGeneratedPageMetadata | undefined>(undefined);
  const [numberingSections, setNumberingSections] = useState<NumberingSection[]>([]);
  const [navigationHistory, setNavigationHistory] = useState<NavigationHistoryState>({ past: [], future: [] });
  const [activeTool, setActiveTool] = useState<Tool>("select");
  const [penColor, setPenColor] = useState("#111827");
  const [penWidth, setPenWidth] = useState(4);
  const [penOpacity, setPenOpacity] = useState(1);
  const [pendingSignature, setPendingSignature] = useState<PreparedImage | null>(null);
  const [copiedMark, setCopiedMark] = useState<Mark | null>(null);
  const [selectedMark, setSelectedMark] = useState<string | null>("demo-title");
  const [commentDraft, setCommentDraft] = useState("");
  const [commentReplyDraft, setCommentReplyDraft] = useState("");
  const [editingReply, setEditingReply] = useState<{ id: string; body: string } | null>(null);
  const [commentFilter, setCommentFilter] = useState<CommentStatusFilter>("open");
  const [showAllComments, setShowAllComments] = useState(false);
  const [editingText, setEditingText] = useState<{ id: string; before: Mark; draft: string; isNew: boolean } | null>(null);
  const [zoom, setZoom] = useState(1);
  const [query, setQuery] = useState("");
  const [leftPanel, setLeftPanel] = useState<LeftPanel>("pages");
  const [activeWorkspace, setActiveWorkspace] = useState<WorkspaceMode>(() => getStoredWorkspaceMode());
  const [isLeftPanelCollapsed, setIsLeftPanelCollapsed] = useState(false);
  const [isRightPanelCollapsed, setIsRightPanelCollapsed] = useState(false);
  const [isShortcutDialogOpen, setIsShortcutDialogOpen] = useState(false);
  const [isCloseGuardOpen, setIsCloseGuardOpen] = useState(false);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const [selectedPageIds, setSelectedPageIds] = useState<string[]>([demoPages[0].id]);
  const [lastSelectedPageId, setLastSelectedPageId] = useState<string | null>(demoPages[0].id);
  const [copiedPages, setCopiedPages] = useState<{ pages: PageView[]; marks: Mark[] } | null>(null);
  const [draggedPageIds, setDraggedPageIds] = useState<string[]>([]);
  const [pageDropTargetId, setPageDropTargetId] = useState<string | null>(null);
  const [pageDropPosition, setPageDropPosition] = useState<"before" | "after">("after");
  const [blankPagePreset, setBlankPagePreset] = useState<BlankPagePreset>("matchCurrent");
  const [blankPageOrientation, setBlankPageOrientation] = useState<PageOrientation>("portrait");
  const [blankPageUnit, setBlankPageUnit] = useState<DimensionUnit>("pt");
  const [blankPageBackground, setBlankPageBackground] = useState("#ffffff");
  const [blankPageQuantity, setBlankPageQuantity] = useState(1);
  const [blankPageLabel, setBlankPageLabel] = useState("");
  const [customPageWidth, setCustomPageWidth] = useState(612);
  const [customPageHeight, setCustomPageHeight] = useState(792);
  const [activePageDialog, setActivePageDialog] = useState<PageDialog>(null);
  const [pageContextMenu, setPageContextMenu] = useState<PageContextMenuState>(null);
  const [mergeQueue, setMergeQueue] = useState<MergeQueueItem[]>([]);
  const [mergeInsertPosition, setMergeInsertPosition] = useState<InsertPosition>("after");
  const [pendingDocxFile, setPendingDocxFile] = useState<File | null>(null);
  const [docxImportOptions, setDocxImportOptions] = useState<DocxImportOptions>(defaultDocxImportOptions);
  const [docxImportMode, setDocxImportMode] = useState<"replace" | "append">("replace");
  const [docxImportReport, setDocxImportReport] = useState<DocxImportReport | null>(null);
  const [docxReimportTarget, setDocxReimportTarget] = useState<DocxImportMetadata | null>(null);
  const [pendingPptxFile, setPendingPptxFile] = useState<File | null>(null);
  const [pptxImportOptions, setPptxImportOptions] = useState<PptxImportOptions>(defaultPptxImportOptions);
  const [pptxImportMode, setPptxImportMode] = useState<"replace" | "append">("replace");
  const [pptxImportReport, setPptxImportReport] = useState<PptxImportReport | null>(null);
  const [pptxReimportTarget, setPptxReimportTarget] = useState<PptxImportMetadata | null>(null);
  const [replaceFile, setReplaceFile] = useState<MergeQueueItem | null>(null);
  const [replaceSourcePage, setReplaceSourcePage] = useState(1);
  const [replacePreserveAnnotations, setReplacePreserveAnnotations] = useState(true);
  const [splitMode, setSplitMode] = useState<SplitMode>("afterCurrent");
  const [splitEveryN, setSplitEveryN] = useState(1);
  const [splitRanges, setSplitRanges] = useState("");
  const [extractDeleteAfter, setExtractDeleteAfter] = useState(false);
  const [extractFilename, setExtractFilename] = useState("");
  const [labelPrefix, setLabelPrefix] = useState("");
  const [labelSuffix, setLabelSuffix] = useState("");
  const [labelStart, setLabelStart] = useState(1);
  const [labelPadding, setLabelPadding] = useState(2);
  const [labelFormat, setLabelFormat] = useState<"number" | "roman" | "alpha">("number");
  const [isDragging, setIsDragging] = useState(false);
  const [workState, setWorkState] = useState<WorkState | null>(null);
  const [errorMessage, setErrorMessage] = useState("");
  const fileInput = useRef<HTMLInputElement>(null);
  const docxInput = useRef<HTMLInputElement>(null);
  const pptxInput = useRef<HTMLInputElement>(null);
  const projectInput = useRef<HTMLInputElement>(null);
  const importPdfInput = useRef<HTMLInputElement>(null);
  const imageInput = useRef<HTMLInputElement>(null);
  const projectImageInput = useRef<HTMLInputElement>(null);
  const projectAssetImportTarget = useRef<ProjectAssetImportTarget | null>(null);
  const signatureInput = useRef<HTMLInputElement>(null);
  const commentEditorRef = useRef<HTMLTextAreaElement>(null);

  const selected = marks.find((mark) => mark.id === selectedMark) ?? null;
  const comments = marks.filter((mark) => mark.kind === "comment" && mark.comment);
  const currentPageView = pages.find((page) => page.pageNumber === currentPage) ?? pages[0];
  const currentPageImportMetadata = currentPageView?.importMetadata;
  const docxImportGroups = useMemo(() => {
    const groups = new Map<string, { metadata: DocxImportMetadata; pages: PageView[]; warnings: number; mappings: number }>();
    for (const page of pages) {
      if (!page.importMetadata || isPptxImportMetadata(page.importMetadata)) continue;
      const existing = groups.get(page.importMetadata.importId) ?? { metadata: page.importMetadata, pages: [], warnings: page.importMetadata.warningCount, mappings: 0 };
      existing.pages.push(page);
      existing.mappings += page.sourceMappings?.length ?? 0;
      groups.set(page.importMetadata.importId, existing);
    }
    return Array.from(groups.values());
  }, [pages]);
  const pptxImportGroups = useMemo(() => {
    const groups = new Map<string, { metadata: PptxImportMetadata; pages: PageView[]; warnings: number; mappings: number; notes: number }>();
    for (const page of pages) {
      if (!page.importMetadata || !isPptxImportMetadata(page.importMetadata)) continue;
      const existing = groups.get(page.importMetadata.importId) ?? { metadata: page.importMetadata, pages: [], warnings: page.importMetadata.warningCount, mappings: 0, notes: page.importMetadata.speakerNotes.length };
      existing.pages.push(page);
      existing.mappings += page.sourceMappings?.length ?? 0;
      groups.set(page.importMetadata.importId, existing);
    }
    return Array.from(groups.values());
  }, [pages]);
  const docxReimportImpact = useMemo<DocxReimportImpact | null>(() => {
    if (!docxReimportTarget) return null;
    const importedPageIds = pages.filter((page) => page.importMetadata?.importId === docxReimportTarget.importId).map((page) => page.id);
    const importedPageIdSet = new Set(importedPageIds);
    return {
      currentPages: importedPageIds.length,
      currentBookmarks: bookmarks.filter((bookmark) => importedPageIdSet.has(bookmark.pageId)).length,
      manualAnnotationsAffected: marks.filter((mark) => mark.pageId && importedPageIdSet.has(mark.pageId) && !mark.comment?.title?.startsWith("Word comment")).length,
      commentsAffected: marks.filter((mark) => mark.kind === "comment" && mark.pageId && importedPageIdSet.has(mark.pageId)).length,
      manualBookmarksPreserved: bookmarks.filter((bookmark) => !importedPageIdSet.has(bookmark.pageId)).length,
      publishingSettingsPreserved: true,
    };
  }, [bookmarks, docxReimportTarget, marks, pages]);
  const pptxReimportImpact = useMemo<PptxReimportImpact | null>(() => {
    if (!pptxReimportTarget) return null;
    const importedPageIds = pages.filter((page) => page.importMetadata?.importId === pptxReimportTarget.importId).map((page) => page.id);
    const importedPageIdSet = new Set(importedPageIds);
    return {
      currentSlides: importedPageIds.length,
      currentBookmarks: bookmarks.filter((bookmark) => importedPageIdSet.has(bookmark.pageId)).length,
      manualAnnotationsAffected: marks.filter((mark) => mark.pageId && importedPageIdSet.has(mark.pageId)).length,
      speakerNotes: pptxReimportTarget.speakerNotes.length,
      publishingSettingsPreserved: true,
    };
  }, [bookmarks, marks, pages, pptxReimportTarget]);
  const selectedPageCount = selectedPageIds.filter((id) => pages.some((page) => page.id === id)).length;
  const activePageIds = selectedPageIds.length > 0 ? selectedPageIds.filter((id) => pages.some((page) => page.id === id)) : currentPageView ? [currentPageView.id] : [];
  const pageAnnotations = useMemo(() => countAnnotationsByPageId(marks, pages), [marks, pages]);
  const publishingReferencedImageAssetIds = useMemo(() => collectPublishingImageAssetIds(publishingSettings), [publishingSettings]);
  const projectAssets = useMemo(() => collectProjectAssets(sourceDocuments, pages, marks, projectImageAssets, publishingReferencedImageAssetIds), [sourceDocuments, pages, marks, projectImageAssets, publishingReferencedImageAssetIds]);
  const publishingImageAssets = useMemo(
    () => {
      const markAssets = marks
        .filter((mark) => (mark.kind === "image" || mark.kind === "pngSignature") && mark.imageDataUrl)
        .map((mark) => ({ id: mark.id, name: mark.imageName || mark.text || mark.kind, dataUrl: mark.imageDataUrl, mimeType: mark.imageMimeType }));
      return [
        ...projectImageAssets.map((asset) => ({ id: asset.id, name: asset.name, dataUrl: asset.dataUrl, mimeType: asset.mimeType })),
        ...markAssets.filter((asset) => !projectImageAssets.some((projectAsset) => projectAsset.id === asset.id)),
      ];
    },
    [marks, projectImageAssets]
  );
  const currentProjectFingerprint = useMemo(
    () => getProjectFingerprint(projectId, projectMetadata, pages, marks, sourceDocuments, pdfName, publishingSettings, projectImageAssets, getNavigationSnapshot()),
    [projectId, projectMetadata, pages, marks, sourceDocuments, pdfName, publishingSettings, projectImageAssets, bookmarks, tocSettings, manualTocEntries, generatedToc, numberingSections]
  );
  const hasUnsavedChanges = hasOpenProject && (isProjectDirty || currentProjectFingerprint !== savedProjectFingerprint);
  const visibleComments = comments.filter((mark) => {
    if (!mark.comment) return false;
    return (showAllComments || mark.page === currentPage) && commentMatchesFilter(mark.comment.resolved, commentFilter);
  });
  const isBusy = workState !== null;
  const canUndo = (history.past.length > 0 || publishingHistory.past.length > 0 || navigationHistory.past.length > 0) && !isBusy;
  const canRedo = (history.future.length > 0 || publishingHistory.future.length > 0 || navigationHistory.future.length > 0) && !isBusy;
  const saveStatus = workState
    ? workState.message
    : hasUnsavedChanges
      ? "Unsaved"
      : lastSavedAt
        ? "Saved"
        : lastAutosavedAt
          ? "Autosaved"
          : hasOpenProject
            ? "Draft"
            : "Welcome";
  const filteredPages = useMemo(() => {
    if (!query.trim()) return pages;
    const normalized = query.trim().toLowerCase();
    return pages.filter((page) => String(page.pageNumber).includes(normalized));
  }, [pages, query]);
  const pageReferences = useMemo<PageReference[]>(
    () => pages.map((page) => ({ id: page.id, pageNumber: page.pageNumber, label: page.label, width: page.width, height: page.height })),
    [pages]
  );
  const visibleBookmarks = useMemo(() => getVisibleBookmarkTree(bookmarks), [bookmarks]);
  const tocEntries = useMemo(() => buildTocEntries(bookmarks, manualTocEntries, pageReferences, tocSettings), [bookmarks, manualTocEntries, pageReferences, tocSettings]);
  const resolvedPageNumbers = useMemo(
    () => resolveNumberingSections(pages, numberingSections, publishingSettings.pageNumbers.customTemplate),
    [pages, numberingSections, publishingSettings.pageNumbers.customTemplate]
  );
  const numberingIssues = useMemo(() => validateNumberingSections(pages, numberingSections), [pages, numberingSections]);
  const numberingSummary = useMemo(() => getNumberingSummary(pages, numberingSections, publishingSettings.pageNumbers.customTemplate), [pages, numberingSections, publishingSettings.pageNumbers.customTemplate]);
  const navigationIssues = useMemo(
    () => validateNavigation({ bookmarks, manualEntries: manualTocEntries, pages: pageReferences, generatedToc }),
    [bookmarks, manualTocEntries, pageReferences, generatedToc]
  );
  const workspaceItems: Array<{ id: WorkspaceMode; icon: ReactElement; title: string; description: string; shortcut: string }> = [
    { id: "import", icon: <FilePlus2 />, title: "Import", description: "Sources and reports", shortcut: "1" },
    { id: "assemble", icon: <Files />, title: "Assemble", description: "Pages and assets", shortcut: "2" },
    { id: "review", icon: <MessageSquare />, title: "Review", description: "Edits, markup, comments", shortcut: "3" },
    { id: "publish", icon: <Download />, title: "Publish", description: "Marks and export", shortcut: "4" },
  ];
  const activeWorkspaceItem = workspaceItems.find((workspace) => workspace.id === activeWorkspace) ?? workspaceItems[0];
  const canShowObjectInspector = activeWorkspace === "review";

  function selectWorkspace(workspace: WorkspaceMode) {
    setActiveWorkspace(workspace);
    window.localStorage.setItem(WORKSPACE_STORAGE_KEY, workspace);
    setIsLeftPanelCollapsed(false);
    if (workspace === "import") setLeftPanel("insert");
    if (workspace === "assemble") setLeftPanel("pages");
    if (workspace === "review") setLeftPanel("comments");
    if (workspace === "publish") setLeftPanel("insert");
  }

  function pushToast(toast: Omit<ToastMessage, "id">) {
    const id = crypto.randomUUID();
    setToasts((existing) => [...existing.filter((item) => item.title !== toast.title || item.detail !== toast.detail).slice(-3), { ...toast, id }]);
    if (toast.tone === "success" || toast.tone === "info") {
      window.setTimeout(() => {
        setToasts((existing) => existing.filter((item) => item.id !== id));
      }, 4200);
    }
  }

  function changeThemeMode(mode: ThemeMode) {
    setThemeMode(mode);
    window.localStorage.setItem(THEME_STORAGE_KEY, mode);
  }

  function getNavigationSnapshot(): NavigationState {
    return {
      bookmarks,
      tocSettings,
      manualTocEntries,
      generatedToc,
      numberingSections,
    };
  }

  function resetNavigationState() {
    setBookmarks([]);
    setSelectedBookmarkId(null);
    setTocSettings(createDefaultTocSettings());
    setManualTocEntries([]);
    setGeneratedToc(undefined);
    setNumberingSections([]);
    setNavigationHistory({ past: [], future: [] });
  }

  function setNavigationFromManifest(navigation?: NavigationState) {
    setBookmarks(normalizeBookmarks(navigation?.bookmarks ?? []));
    setSelectedBookmarkId(null);
    setTocSettings({ ...createDefaultTocSettings(), ...(navigation?.tocSettings ?? {}) });
    setManualTocEntries(navigation?.manualTocEntries ?? []);
    setGeneratedToc(navigation?.generatedToc);
    setNumberingSections(navigation?.numberingSections ?? []);
    setNavigationHistory({ past: [], future: [] });
  }

  function commitNavigation(next: Partial<NavigationState>, markTocStale = true) {
    const before = getNavigationSnapshot();
    const hasGeneratedTocPatch = Object.prototype.hasOwnProperty.call(next, "generatedToc");
    const after: NavigationState = {
      bookmarks: next.bookmarks ?? bookmarks,
      tocSettings: next.tocSettings ?? tocSettings,
      manualTocEntries: next.manualTocEntries ?? manualTocEntries,
      numberingSections: next.numberingSections ?? numberingSections,
      generatedToc: hasGeneratedTocPatch ? next.generatedToc : markTocStale && generatedToc ? { ...generatedToc, stale: true } : generatedToc,
    };
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    setNavigationHistory((existing) => ({
      past: [...existing.past.slice(-(MAX_HISTORY_ENTRIES - 1)), before],
      future: [],
    }));
    setBookmarks(normalizeBookmarks(after.bookmarks));
    setTocSettings(after.tocSettings);
    setManualTocEntries(after.manualTocEntries);
    setGeneratedToc(after.generatedToc);
    setNumberingSections(after.numberingSections ?? []);
    setIsProjectDirty(true);
  }

  function getProjectSnapshotInput() {
    return {
      projectId,
      metadata: projectMetadata,
      pages,
      annotations: marks,
      projectImageAssets,
      publishingSettings,
      navigation: getNavigationSnapshot(),
      sourceDocuments,
      workspaceState: {
        currentPage,
        zoom,
        activeWorkspace,
        leftPanel,
        selectedPageIds,
      },
      exportSettings: {
        defaultFileName: pdfName.replace(/\.pdf$/i, "") || projectMetadata.name,
      },
    };
  }

  function createProjectFileBytes() {
    return serializeProject(getProjectSnapshotInput());
  }

  async function loadStoredAutosave() {
    return runtime.isDesktop ? runtime.loadAutosave() : loadBrowserAutosave();
  }

  async function saveStoredAutosave(bytes: Uint8Array, manifest: ReturnType<typeof buildProjectManifest>) {
    return runtime.isDesktop ? runtime.saveAutosave(bytes, manifest, desktopProjectPath) : saveBrowserAutosave(bytes, manifest);
  }

  async function clearStoredAutosave() {
    if (runtime.isDesktop) {
      await runtime.clearAutosave();
      return;
    }
    await clearBrowserAutosave();
  }

  async function addStoredRecentProject(manifest: ReturnType<typeof buildProjectManifest>, options: { filename?: string; path?: string; origin?: RecentProject["origin"]; autosaveAvailable?: boolean } = {}) {
    if (runtime.isDesktop) {
      await runtime.addRecentProject(manifest, options);
      setRecentProjects((await runtime.loadRecentProjects()) ?? []);
      return;
    }
    await addBrowserRecentProject(manifest, { filename: options.filename, origin: options.origin, autosaveAvailable: options.autosaveAvailable });
    setRecentProjects(loadBrowserRecentProjects());
  }

  async function refreshStoredRecentProjects() {
    setRecentProjects(runtime.isDesktop ? (await runtime.loadRecentProjects()) ?? [] : loadBrowserRecentProjects());
  }

  function startNewProject() {
    setNewProjectDraft(createProjectMetadata("Untitled project"));
    setIsProjectDialogOpen(true);
  }

  function createNewProjectFromDraft(mode: "blank" | "pdf") {
    if (hasOpenProject && hasUnsavedChanges && !window.confirm("Close the current project and discard unsaved changes?")) return;
    if (mode === "pdf") {
      setIsProjectDialogOpen(false);
      void choosePdfFile();
      return;
    }
    const name = newProjectDraft.name.trim() || "Untitled project";
    const blankPage = createBlankPageView(1, { width: 612, height: 792 });
    setHasOpenProject(true);
    setProjectId(crypto.randomUUID());
    setProjectMetadata({ ...newProjectDraft, name, modifiedAt: new Date().toISOString() });
    setSavedProjectFingerprint("");
    setLastSavedAt(null);
    setLastAutosavedAt(null);
    setDesktopProjectPath(null);
    setPdfName(name);
    setPdfBytes(null);
    setPdfDoc(null);
    setActiveSourceDocumentId(null);
    setSourceDocuments({});
    setProjectImageAssets([]);
    setPages([blankPage]);
    setMarks([]);
    resetNavigationState();
    setPublishingSettings(createDefaultPublishingSettings());
    setPublishingHistory({ past: [], future: [] });
    setHistory({ past: [], future: [] });
    setSelectedMark(null);
    setSelectedPageIds([blankPage.id]);
    setLastSelectedPageId(blankPage.id);
    setCurrentPage(1);
    setActiveWorkspace("assemble");
    setIsProjectDirty(true);
    setIsProjectDialogOpen(false);
    setProjectStatusMessage("New project created. Autosave will keep a local recovery copy.");
    setErrorMessage("");
  }

  async function openProjectFile(file: File, sourcePath?: string) {
    try {
      setErrorMessage("");
      setWorkState({ message: `Opening ${file.name}`, progress: 15 });
      const bytes = new Uint8Array(await file.arrayBuffer());
      await loadProjectBytes(bytes, file.name, sourcePath);
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(getProjectErrorMessage(error, "open"));
    }
  }

  async function loadProjectBytes(bytes: Uint8Array, fileName = "Project.pproj", sourcePath?: string) {
    const bundle = deserializeProject(bytes);
    const nextSources: Record<string, SourceDocument> = {};
    for (const source of bundle.manifest.sources) {
      const fileBytes = bundle.files[source.path];
      if (!fileBytes) throw new Error(`Project is missing source asset: ${source.name}`);
      nextSources[source.id] = { id: source.id, name: source.name, bytes: fileBytes, mimeType: source.mimeType };
    }

    const nextPages = renumberPages(bundle.manifest.pages as PageView[]);
    const nextMarks = migrateMarksToPageIds(bundle.manifest.annotations as Mark[], nextPages);
    const syncedMarks = syncMarksToPages(nextMarks, nextPages);
    const markAssetIds = new Set(syncedMarks.filter((mark) => (mark.kind === "image" || mark.kind === "pngSignature") && mark.imageDataUrl).map((mark) => mark.id));
    const nextProjectImageAssets = bundle.manifest.assets.reduce<ProjectImageAsset[]>((assets, asset) => {
      if (asset.type !== "image" || markAssetIds.has(asset.id)) return assets;
      const fileBytes = bundle.files[asset.path];
      if (!fileBytes) return assets;
      assets.push({
        id: asset.id,
        name: asset.name,
        dataUrl: bytesToDataUrl(fileBytes, asset.mimeType === "image/jpeg" ? "image/jpeg" : "image/png"),
        mimeType: asset.mimeType === "image/jpeg" ? "image/jpeg" : "image/png",
        naturalWidth: 0,
        naturalHeight: 0,
        createdAt: asset.dateAdded,
        updatedAt: bundle.manifest.metadata.modifiedAt,
      });
      return assets;
    }, []);
    const firstSource = Object.values(nextSources).find((source) => (source.mimeType ?? "application/pdf") === "application/pdf") ?? Object.values(nextSources)[0];
    const nextNavigation: NavigationState = {
      bookmarks: normalizeBookmarks(bundle.manifest.navigation?.bookmarks ?? []),
      tocSettings: { ...createDefaultTocSettings(), ...(bundle.manifest.navigation?.tocSettings ?? {}) },
      manualTocEntries: bundle.manifest.navigation?.manualTocEntries ?? [],
      generatedToc: bundle.manifest.navigation?.generatedToc,
      numberingSections: bundle.manifest.navigation?.numberingSections ?? [],
    };
    setHasOpenProject(true);
    setProjectId(bundle.manifest.projectId);
    setProjectMetadata(bundle.manifest.metadata);
    const nextPublishingSettings = mergePublishingSettings(bundle.manifest.publishingSettings);
    setSavedProjectFingerprint(getProjectFingerprint(bundle.manifest.projectId, bundle.manifest.metadata, nextPages, syncedMarks, nextSources, bundle.manifest.exportSettings.defaultFileName || fileName.replace(/\.pproj$/i, ""), nextPublishingSettings, nextProjectImageAssets, nextNavigation));
    setPdfName(bundle.manifest.exportSettings.defaultFileName || fileName.replace(/\.pproj$/i, ""));
    setPdfBytes(firstSource?.bytes ?? null);
    setPdfDoc(null);
    setActiveSourceDocumentId(firstSource?.id ?? null);
    setSourceDocuments(nextSources);
    setProjectImageAssets(nextProjectImageAssets);
    setPages(nextPages);
    setMarks(syncedMarks);
    setNavigationFromManifest(nextNavigation);
    setPublishingSettings(nextPublishingSettings);
    setPublishingHistory({ past: [], future: [] });
    setHistory({ past: [], future: [] });
    setSelectedMark(null);
    setCurrentPage(Math.min(Math.max(1, bundle.manifest.workspaceState.currentPage || 1), Math.max(1, nextPages.length)));
    setZoom(bundle.manifest.workspaceState.zoom || 1);
    setActiveWorkspace(normalizeWorkspaceMode(bundle.manifest.workspaceState.activeWorkspace));
    setLeftPanel((bundle.manifest.workspaceState.leftPanel as LeftPanel) || "pages");
    const selectedIds = bundle.manifest.workspaceState.selectedPageIds.filter((id) => nextPages.some((page) => page.id === id));
    setSelectedPageIds(selectedIds.length > 0 ? selectedIds : nextPages[0] ? [nextPages[0].id] : []);
    setLastSelectedPageId(selectedIds[0] ?? nextPages[0]?.id ?? null);
    setIsProjectDirty(false);
    setLastSavedAt(bundle.manifest.metadata.modifiedAt);
    setDesktopProjectPath(sourcePath ?? null);
    setProjectStatusMessage(`Opened ${fileName}.`);
    await addStoredRecentProject(bundle.manifest, { filename: fileName, path: sourcePath, origin: runtime.isDesktop ? "file-system-access" : "browser-import" });
  }

  function handleProjectUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    if (hasOpenProject && hasUnsavedChanges && !window.confirm("Open this project and discard unsaved changes in the current one?")) return;
    void openProjectFile(file);
  }

  async function chooseProjectFile() {
    if (!runtime.isDesktop) {
      projectInput.current?.click();
      return;
    }
    const files = await runtime.openFiles({ kind: "project", title: "Open Publish Pro Project" });
    const file = files?.[0];
    if (!file) return;
    if (hasOpenProject && hasUnsavedChanges && !window.confirm("Open this project and discard unsaved changes in the current one?")) return;
    await openProjectFile(await runtimeFileToBrowserFile(file), file.path);
  }

  async function choosePdfFile() {
    if (!runtime.isDesktop) {
      fileInput.current?.click();
      return;
    }
    const files = await runtime.openFiles({ kind: "pdf", title: "Open PDF" });
    const file = files?.[0];
    if (!file) return;
    if (hasOpenProject && hasUnsavedChanges && !window.confirm("Open this PDF and discard unsaved changes in the current project?")) return;
    await loadFile(await runtimeFileToBrowserFile(file));
  }

  async function chooseDocxFile() {
    if (!runtime.isDesktop) {
      docxInput.current?.click();
      return;
    }
    const files = await runtime.openFiles({ kind: "docx", title: "Import Word Document" });
    const file = files?.[0];
    if (!file) return;
    beginDocxImport(await runtimeFileToBrowserFile(file), hasOpenProject ? "append" : "replace");
  }

  async function choosePptxFile() {
    if (!runtime.isDesktop) {
      pptxInput.current?.click();
      return;
    }
    const files = await runtime.openFiles({ kind: "pptx", title: "Import PowerPoint Presentation" });
    const file = files?.[0];
    if (!file) return;
    beginPptxImport(await runtimeFileToBrowserFile(file), hasOpenProject ? "append" : "replace");
  }

  async function choosePdfImports() {
    if (!runtime.isDesktop) {
      importPdfInput.current?.click();
      return;
    }
    const files = await runtime.openFiles({ kind: "pdf", multiple: true, title: "Import PDF Files" });
    if (!files?.length) return;
    await addFilesToMergeQueue(await Promise.all(files.map(runtimeFileToBrowserFile)));
    setActivePageDialog("merge");
  }

  async function saveProject(targetPath = desktopProjectPath) {
    try {
      setErrorMessage("");
      const bytes = createProjectFileBytes();
      const manifest = buildProjectManifest(getProjectSnapshotInput());
      const filename = `${sanitizeFilename(projectMetadata.name || "publish-pro-project")}.pproj`;
      const desktopResult = runtime.isDesktop ? await runtime.saveFile({ kind: "project", bytes, suggestedName: filename, currentPath: targetPath }) : null;
      if (desktopResult && !desktopResult.saved) return false;
      const savedViaFileSystem = desktopResult?.native ?? await saveProjectBytes(bytes, filename);
      if (desktopResult?.path) setDesktopProjectPath(desktopResult.path);
      await addStoredRecentProject(manifest, { filename, path: desktopResult?.path ?? targetPath ?? undefined, origin: savedViaFileSystem ? "file-system-access" : "download" });
      await clearStoredAutosave();
      setAutosaveCandidate(null);
      setLastSavedAt(new Date().toISOString());
      setProjectStatusMessage(runtime.isDesktop && desktopResult?.path ? "Project saved to disk." : savedViaFileSystem ? "Project saved." : "Project downloaded as a .pproj file.");
      setIsProjectDirty(false);
      setSavedProjectFingerprint(currentProjectFingerprint);
      return true;
    } catch (error) {
      setErrorMessage(getProjectErrorMessage(error, "save"));
      return false;
    }
  }

  function saveProjectAs() {
    void saveProject(null);
  }

  function closeProject() {
    if (hasOpenProject && hasUnsavedChanges && !window.confirm("Close this project and discard unsaved changes?")) return;
    setHasOpenProject(false);
    setPdfName("Untitled document");
    setDesktopProjectPath(null);
    setPdfBytes(null);
    setPdfDoc(null);
    setActiveSourceDocumentId(null);
    setSourceDocuments({});
    setProjectImageAssets([]);
    setPages(demoPages);
    setMarks(initialMarks);
    resetNavigationState();
    setPublishingSettings(createDefaultPublishingSettings());
    setPublishingHistory({ past: [], future: [] });
    setHistory({ past: [], future: [] });
    setSelectedMark("demo-title");
    setSelectedPageIds([demoPages[0].id]);
    setLastSelectedPageId(demoPages[0].id);
    setCurrentPage(1);
    setIsProjectDirty(false);
    setSavedProjectFingerprint("");
    setLastSavedAt(null);
    setLastAutosavedAt(null);
    setProjectStatusMessage("");
    setErrorMessage("");
  }

  function restoreAutosave() {
    if (!autosaveCandidate) return;
    if (hasOpenProject && hasUnsavedChanges && !window.confirm("Restore autosave and discard unsaved changes in the current project?")) return;
    try {
      void loadProjectBytes(autosaveCandidate.bytes, `${autosaveCandidate.manifest.metadata.name}.pproj`);
      setProjectStatusMessage(`Restored autosave from ${formatDateTime(autosaveCandidate.savedAt)}.`);
    } catch (error) {
      setErrorMessage(getProjectErrorMessage(error, "restore"));
    }
  }

  function reopenRecentProject(project: RecentProject) {
    setErrorMessage(`${project.name} is listed in Recents. Browser security does not allow Publish Pro to reopen a downloaded project file automatically; choose Open Project and select ${project.savedFilename ?? "the .pproj file"}.`);
  }

  function updateProjectMetadata(patch: Partial<ProjectMetadata>) {
    setProjectMetadata((metadata) => ({ ...metadata, ...patch, modifiedAt: new Date().toISOString() }));
    if (patch.name) setPdfName(patch.name);
    setIsProjectDirty(true);
  }

  function updatePublishingSettings(updater: (settings: PublishingSettings) => PublishingSettings, recordHistory = true) {
    setPublishingSettings((existing) => {
      const next = mergePublishingSettings(updater(existing));
      if (JSON.stringify(existing) === JSON.stringify(next)) return existing;
      if (recordHistory) {
        setPublishingHistory((historyState) => ({
          past: [...historyState.past.slice(-49), existing],
          future: [],
        }));
      }
      setIsProjectDirty(true);
      setPublishingValidationMessage(validatePublishingSettings(next));
      return next;
    });
  }

  function applyPublishingPreset(preset: PublishingPreset) {
    updatePublishingSettings((settings) => mergePublishingSettings({ ...settings, ...preset.settings }));
  }

  function saveCurrentPublishingPreset(type: PublishingPreset["type"]) {
    const name = window.prompt("Preset name", type === "watermark" ? "Watermark preset" : type === "pageNumbers" ? "Page numbering preset" : "Header footer preset")?.trim();
    if (!name) return;
    const preset: PublishingPreset = {
      id: crypto.randomUUID(),
      name,
      type,
      createdAt: new Date().toISOString(),
      settings:
        type === "watermark"
          ? { watermark: publishingSettings.watermark }
          : type === "pageNumbers"
          ? { pageNumbers: publishingSettings.pageNumbers }
          : { headerFooter: publishingSettings.headerFooter },
    };
    savePublishingPreset(preset);
    setPublishingPresets(loadPublishingPresets());
  }

  function renamePublishingPreset(preset: PublishingPreset) {
    const name = window.prompt("Preset name", preset.name)?.trim();
    if (!name || name === preset.name) return;
    savePublishingPreset({ ...preset, name });
    setPublishingPresets(loadPublishingPresets());
  }

  function updateHeaderFooterZone(area: "header" | "footer", side: "left" | "center" | "right", patch: Partial<HeaderFooterZone>) {
    updatePublishingSettings((settings) => ({
      ...settings,
      headerFooter: {
        ...settings.headerFooter,
        [area]: {
          ...settings.headerFooter[area],
          [side]: {
            ...settings.headerFooter[area][side],
            ...patch,
            image: Object.prototype.hasOwnProperty.call(patch, "image") ? patch.image : settings.headerFooter[area][side].image,
          },
        },
      },
    }), patch.text === undefined);
  }

  function updateHeaderFooterZoneImage(area: "header" | "footer", side: "left" | "center" | "right", patch: Partial<HeaderFooterZoneImage>) {
    updateHeaderFooterZone(area, side, {
      image: {
        ...createDefaultHeaderFooterImage(),
        ...publishingSettings.headerFooter[area][side].image,
        ...patch,
      },
    });
  }

  function updateHeaderFooterOverrideZone(key: "firstPage" | "oddPage" | "evenPage", zone: PublishingZone, patch: Partial<HeaderFooterZone>) {
    updatePublishingSettings((settings) => {
      const existing = mergeHeaderFooterZone(settings.headerFooter[key]?.[zone]);
      return {
        ...settings,
        headerFooter: {
          ...settings.headerFooter,
          [key]: {
            ...(settings.headerFooter[key] ?? {}),
            [zone]: {
              ...existing,
              ...patch,
              image: Object.prototype.hasOwnProperty.call(patch, "image") ? patch.image : existing.image,
            },
          },
        },
      };
    }, patch.text === undefined);
  }

  function updateHeaderFooterOverrideImage(key: "firstPage" | "oddPage" | "evenPage", zone: PublishingZone, patch: Partial<HeaderFooterZoneImage>) {
    const existing = mergeHeaderFooterZone(publishingSettings.headerFooter[key]?.[zone]);
    updateHeaderFooterOverrideZone(key, zone, {
      image: {
        ...createDefaultHeaderFooterImage(),
        ...existing.image,
        ...patch,
      },
    });
  }

  function locateProjectAsset(assetId: string) {
    const sourcePage = pages.find((page) => page.sourceDocumentId === assetId);
    if (sourcePage) {
      setCurrentPage(sourcePage.pageNumber);
      setSelectedPageIds([sourcePage.id]);
      setLastSelectedPageId(sourcePage.id);
      setActiveWorkspace("assemble");
      return;
    }
    const mark = marks.find((item) => item.id === assetId);
    if (mark) {
      setCurrentPage(mark.page);
      setSelectedMark(mark.id);
      setActiveWorkspace("review");
      return;
    }
    if (projectImageAssets.some((asset) => asset.id === assetId)) {
      setActiveWorkspace("assemble");
      setLeftPanel("insert");
    }
  }

  function renameProjectAsset(assetId: string) {
    const asset = projectAssets.find((item) => item.id === assetId);
    if (!asset) return;
    const nextName = window.prompt("Asset name", asset.name)?.trim();
    if (!nextName || nextName === asset.name) return;
    if (asset.type === "source-pdf" || asset.type === "source-docx" || asset.type === "source-pptx") {
      setSourceDocuments((existing) => ({
        ...existing,
        ...(existing[assetId] ? { [assetId]: { ...existing[assetId], name: nextName } } : {}),
      }));
    } else if (projectImageAssets.some((item) => item.id === assetId)) {
      setProjectImageAssets((existing) => existing.map((item) => (item.id === assetId ? { ...item, name: nextName, updatedAt: new Date().toISOString() } : item)));
    } else {
      setMarks((existing) => existing.map((mark) => (mark.id === assetId ? { ...mark, imageName: nextName, text: nextName } : mark)));
    }
    setIsProjectDirty(true);
  }

  function removeUnusedAsset(assetId: string) {
    const asset = projectAssets.find((item) => item.id === assetId);
    if (!asset) return;
    if (asset.usageCount > 0) {
      setErrorMessage(`${asset.name} is currently used in this project. Remove or replace the object before deleting the asset.`);
      return;
    }
    if (asset.type === "source-pdf" || asset.type === "source-docx" || asset.type === "source-pptx") {
      setSourceDocuments((existing) => {
        const next = { ...existing };
        delete next[assetId];
        return next;
      });
      setIsProjectDirty(true);
    } else if (projectImageAssets.some((item) => item.id === assetId)) {
      setProjectImageAssets((existing) => existing.filter((item) => item.id !== assetId));
      setIsProjectDirty(true);
    }
  }

  useEffect(() => {
    const mediaQuery = window.matchMedia("(prefers-color-scheme: dark)");

    function applyTheme() {
      const nextResolvedTheme = resolveThemeMode(themeMode);
      setResolvedTheme(nextResolvedTheme);
      document.documentElement.dataset.themePreference = themeMode;
      document.documentElement.dataset.theme = nextResolvedTheme;
      document.documentElement.style.colorScheme = nextResolvedTheme;
    }

    applyTheme();
    if (themeMode !== "system") return undefined;
    mediaQuery.addEventListener("change", applyTheme);
    return () => mediaQuery.removeEventListener("change", applyTheme);
  }, [themeMode]);

  useEffect(() => {
    let cancelled = false;
    refreshStoredRecentProjects().catch(() => undefined);
    loadStoredAutosave()
      .then((autosave) => {
        if (cancelled) return;
        setAutosaveCandidate(autosave);
        setLastAutosavedAt(autosave?.savedAt ?? null);
      })
      .catch(() => {
        if (!cancelled) setProjectStatusMessage("Autosave storage is unavailable in this browser session.");
      });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!hasOpenProject) return undefined;
    const timeout = window.setTimeout(() => {
      void (async () => {
        try {
          setProjectStatusMessage("Autosaving locally...");
          const bytes = createProjectFileBytes();
          const manifest = buildProjectManifest(getProjectSnapshotInput());
          const autosave = await saveStoredAutosave(bytes, manifest);
          await addStoredRecentProject(manifest, { origin: "autosave", autosaveAvailable: true, path: desktopProjectPath ?? undefined });
          setAutosaveCandidate(autosave);
          setLastAutosavedAt(autosave?.savedAt ?? null);
          setProjectStatusMessage(autosave ? `Autosaved locally at ${formatDateTime(autosave.savedAt)}.` : "Autosave is handled by the active runtime.");
        } catch {
          setProjectStatusMessage("Autosave failed. Your current browser session is still active.");
        }
      })();
    }, 900);

    return () => window.clearTimeout(timeout);
  }, [hasOpenProject, projectId, projectMetadata, pages, marks, sourceDocuments, currentPage, zoom, activeWorkspace, leftPanel, selectedPageIds, pdfName, desktopProjectPath]);

  useEffect(() => {
    if (!projectStatusMessage) return;
    pushToast({ tone: projectStatusMessage.toLowerCase().includes("failed") ? "warning" : "info", title: projectStatusMessage });
  }, [projectStatusMessage]);

  useEffect(() => {
    if (!errorMessage) return;
    pushToast({ tone: "error", title: "Action failed", detail: errorMessage });
  }, [errorMessage]);

  useEffect(() => {
    if (!runtime.isDesktop) return;
    const projectTitle = hasOpenProject ? projectMetadata.name || pdfName.replace(/\.pdf$/i, "") || "Untitled Project" : "";
    const title = projectTitle ? `${projectTitle}${hasUnsavedChanges ? " • Unsaved" : ""} - Publish Pro` : "Publish Pro";
    void runtime.setWindowTitle(title);
  }, [hasOpenProject, hasUnsavedChanges, projectMetadata.name, pdfName]);

  useEffect(() => {
    if (!runtime.isDesktop) return undefined;
    let unlistenMenu: (() => void) | undefined;
    let unlistenLaunch: (() => void) | undefined;
    void runtime.listenMenuActions((action) => {
      void handleDesktopMenuAction(action);
    }).then((unlisten) => {
      unlistenMenu = unlisten;
    });
    void runtime.listenLaunchFiles((paths) => {
      void handleDesktopLaunchFiles(paths);
    }).then((unlisten) => {
      unlistenLaunch = unlisten;
    });
    return () => {
      unlistenMenu?.();
      unlistenLaunch?.();
    };
  }, [hasOpenProject, hasUnsavedChanges, desktopProjectPath, pages, marks, activeWorkspace, selected]);

  async function handleDesktopMenuAction(action: string) {
    if (action === "window-close-requested" || action === "exit") {
      if (hasOpenProject && hasUnsavedChanges) {
        setIsCloseGuardOpen(true);
      } else {
        await runtime.forceClose();
      }
      return;
    }
    if (action === "new-project") startNewProject();
    if (action === "open-project") await chooseProjectFile();
    if (action === "save-project") await saveProject();
    if (action === "save-project-as") saveProjectAs();
    if (action === "import-pdf" || action === "open-pdf") await choosePdfFile();
    if (action === "import-docx") await chooseDocxFile();
    if (action === "import-pptx") await choosePptxFile();
    if (action === "publish-pdf") await exportPdf();
    if (action === "close-project") closeProject();
    if (action === "undo") undo();
    if (action === "redo") redo();
    if (action === "delete") removeSelectedMark();
    if (action === "select-all-pages") selectAllPages();
    if (action === "workspace-import") selectWorkspace("import");
    if (action === "workspace-assemble") selectWorkspace("assemble");
    if (action === "workspace-review") selectWorkspace("review");
    if (action === "workspace-publish") selectWorkspace("publish");
    if (action === "zoom-in") setZoom((value) => Math.min(1.6, value + 0.1));
    if (action === "zoom-out") setZoom((value) => Math.max(0.55, value - 0.1));
    if (action === "fit-page") setZoom(1);
    if (action === "fit-width") setZoom(1.25);
    if (action === "toggle-left-panel") setIsLeftPanelCollapsed((value) => !value);
    if (action === "toggle-inspector") setIsRightPanelCollapsed((value) => !value);
    if (action === "shortcuts") setIsShortcutDialogOpen(true);
    if (action === "about") pushToast({ tone: "info", title: "Publish Pro 0.1.0", detail: "By Designovation. Local-first desktop publishing workspace." });
  }

  async function saveAndCloseDesktop() {
    if (await saveProject()) await runtime.forceClose();
  }

  async function discardAndCloseDesktop() {
    setIsCloseGuardOpen(false);
    await runtime.forceClose();
  }

  async function handleDesktopLaunchFiles(paths: string[]) {
    const first = paths[0];
    if (!first) return;
    if (hasOpenProject && hasUnsavedChanges) {
      pushToast({ tone: "warning", title: "Open file blocked", detail: "Save or close the current project before opening a launch file." });
      return;
    }
    const runtimeFile = await runtime.readPath(first);
    if (!runtimeFile) return;
    const file = await runtimeFileToBrowserFile(runtimeFile);
    const name = runtimeFile.name.toLowerCase();
    if (name.endsWith(".pproj")) {
      await openProjectFile(file, runtimeFile.path);
    } else if (name.endsWith(".pdf")) {
      await loadFile(file);
    } else if (name.endsWith(".docx")) {
      beginDocxImport(file, hasOpenProject ? "append" : "replace");
    } else if (name.endsWith(".pptx")) {
      beginPptxImport(file, hasOpenProject ? "append" : "replace");
    } else {
      pushToast({ tone: "warning", title: "Unsupported file", detail: runtimeFile.name });
    }
  }

  function getPageByNumber(pageNumber: number) {
    return pages.find((page) => page.pageNumber === pageNumber) ?? pages[0];
  }

  function getPageIdForNumber(pageNumber: number) {
    return getPageByNumber(pageNumber)?.id;
  }

  function markPageId(mark: Mark, pageSet = pages) {
    if (mark.pageId && pageSet.some((page) => page.id === mark.pageId)) return mark.pageId;
    return pageSet.find((page) => page.pageNumber === mark.page)?.id;
  }

  function markBelongsToPage(mark: Mark, page: PageView) {
    return markPageId(mark) === page.id;
  }

  function marksForPageIds(pageIds: string[], markSet = marks, pageSet = pages) {
    const idSet = new Set(pageIds);
    return markSet.filter((mark) => {
      const id = markPageId(mark, pageSet);
      return id ? idSet.has(id) : false;
    });
  }

  function withPageId(mark: Mark, pageNumber = mark.page): Mark {
    return { ...mark, pageId: mark.pageId ?? getPageIdForNumber(pageNumber), page: pageNumber };
  }

  function getPublishingTokenContext() {
    return {
      projectName: projectMetadata.name || pdfName.replace(/\.pdf$/i, ""),
      client: projectMetadata.client,
      filename: pdfName,
      date: new Intl.DateTimeFormat(undefined, { dateStyle: "medium" }).format(new Date()),
    };
  }

  useEffect(() => {
    function handleKeyboardShortcut(event: KeyboardEvent) {
      const isModifierPressed = event.metaKey || event.ctrlKey;
      const key = event.key.toLowerCase();
      if (!isModifierPressed && key === "?" && !isEditableElement(event.target)) {
        event.preventDefault();
        setIsShortcutDialogOpen(true);
        return;
      }
      if (!isModifierPressed || isBusy) return;

      if (isEditableElement(event.target) && (key === "z" || key === "y" || key === "c" || key === "v")) return;
      if (!isEditableElement(event.target) && ["1", "2", "3", "4"].includes(key)) {
        event.preventDefault();
        selectWorkspace(key === "1" ? "import" : key === "2" ? "assemble" : key === "3" ? "review" : "publish");
        return;
      }
      if (!isEditableElement(event.target) && key === "n") {
        event.preventDefault();
        startNewProject();
        return;
      }
      if (!isEditableElement(event.target) && key === "o") {
        event.preventDefault();
        void chooseProjectFile();
        return;
      }
      if (!isEditableElement(event.target) && key === "s") {
        event.preventDefault();
        if (event.shiftKey) {
          saveProjectAs();
        } else if (hasOpenProject) {
          void saveProject();
        }
        return;
      }
      if (!isEditableElement(event.target) && key === "f") {
        event.preventDefault();
        setLeftPanel("search");
        setIsLeftPanelCollapsed(false);
        return;
      }
      if (!isEditableElement(event.target) && key === "k") {
        event.preventDefault();
        setIsShortcutDialogOpen(true);
        return;
      }
      if (!isEditableElement(event.target) && key === "p" && event.shiftKey) {
        event.preventDefault();
        selectWorkspace("publish");
        return;
      }
      const isUndo = key === "z" && !event.shiftKey;
      const isRedo = (key === "z" && event.shiftKey) || key === "y";
      const isCopy = key === "c" && !event.shiftKey && !isEditableElement(event.target);
      const isPaste = key === "v" && !event.shiftKey && !isEditableElement(event.target);
      const isDuplicate = key === "d" && !event.shiftKey && !isEditableElement(event.target);

      if (isUndo && history.past.length > 0) {
        event.preventDefault();
        undo();
      }

      if (isRedo && history.future.length > 0) {
        event.preventDefault();
        redo();
      }

      if (isCopy && selected) {
        event.preventDefault();
        setCopiedMark(selected);
      }

      if (isPaste && copiedMark) {
        event.preventDefault();
        pasteCopiedMark();
      }

      if (isDuplicate && selected) {
        event.preventDefault();
        duplicateSelectedMark();
      }
    }

    window.addEventListener("keydown", handleKeyboardShortcut);

    return () => {
      window.removeEventListener("keydown", handleKeyboardShortcut);
    };
  }, [history, isBusy, selected, copiedMark, currentPage, marks, pages, hasOpenProject, hasUnsavedChanges]);

  useEffect(() => {
    function handleDeleteShortcut(event: KeyboardEvent) {
      if (isBusy || !selectedMark) return;
      if (event.key !== "Delete" && event.key !== "Backspace") return;
      if (isEditableElement(event.target)) return;

      event.preventDefault();
      removeSelectedMark();
    }

    window.addEventListener("keydown", handleDeleteShortcut);

    return () => {
      window.removeEventListener("keydown", handleDeleteShortcut);
    };
  }, [isBusy, selectedMark, marks]);

  useEffect(() => {
    function handleEscapeShortcut(event: KeyboardEvent) {
      if (isBusy || event.key !== "Escape" || isEditableElement(event.target)) return;
      if (activeTool !== "select") {
        event.preventDefault();
        setActiveTool("select");
        return;
      }
      if (selectedMark) {
        event.preventDefault();
        setSelectedMark(null);
      }
    }

    window.addEventListener("keydown", handleEscapeShortcut);

    return () => {
      window.removeEventListener("keydown", handleEscapeShortcut);
    };
  }, [activeTool, isBusy, selectedMark]);

  useEffect(() => {
    if (selected?.kind !== "comment" || !selected.comment) {
      setCommentDraft("");
      setCommentReplyDraft("");
      setEditingReply(null);
      return;
    }

    setCommentDraft(selected.comment.body);
    setCommentReplyDraft("");
    setEditingReply(null);
    window.setTimeout(() => commentEditorRef.current?.focus(), 0);
  }, [selectedMark]);

  useEffect(() => {
    function handleEscape(event: KeyboardEvent) {
      if (event.key !== "Escape") return;
      if (selected?.kind === "comment") {
        event.preventDefault();
        cancelCommentEdit();
        setSelectedMark(null);
        return;
      }
      if (!isTextMarkupKind(activeTool)) return;

      event.preventDefault();
      window.getSelection()?.removeAllRanges();
      setActiveTool("select");
    }

    window.addEventListener("keydown", handleEscape);

    return () => {
      window.removeEventListener("keydown", handleEscape);
    };
  }, [activeTool, selected]);

  useEffect(() => {
    if (!pageContextMenu && !activePageDialog) return;
    function handleDismiss(event: KeyboardEvent | MouseEvent) {
      if (event instanceof KeyboardEvent && event.key !== "Escape") return;
      if (event instanceof MouseEvent && (event.target as HTMLElement).closest(".page-context-menu, .modal-card")) return;
      if (event instanceof KeyboardEvent) event.preventDefault();
      setPageContextMenu(null);
      if (event instanceof KeyboardEvent) setActivePageDialog(null);
    }

    window.addEventListener("keydown", handleDismiss);
    window.addEventListener("mousedown", handleDismiss);
    return () => {
      window.removeEventListener("keydown", handleDismiss);
      window.removeEventListener("mousedown", handleDismiss);
    };
  }, [pageContextMenu, activePageDialog]);

  useEffect(() => {
    if (!pdfDoc) return;

    let cancelled = false;
    const activePdf = pdfDoc;

    async function renderPages() {
      try {
        setErrorMessage("");
        const nextPages: PageView[] = [];
        for (let index = 1; index <= activePdf.numPages; index += 1) {
          setWorkState({
            message: `Rendering page ${index} of ${activePdf.numPages}`,
            progress: ((index - 1) / activePdf.numPages) * 100,
          });
          const page = await activePdf.getPage(index);
          const rendered = await renderPage(page, activeSourceDocumentId ?? undefined);
          if (cancelled) return;
          nextPages.push(rendered);
          setWorkState({
            message: `Rendering page ${index} of ${activePdf.numPages}`,
            progress: (index / activePdf.numPages) * 100,
          });
        }
        setPages(nextPages);
        setCurrentPage(1);
        setSelectedPageIds(nextPages[0] ? [nextPages[0].id] : []);
        setLastSelectedPageId(nextPages[0]?.id ?? null);
        const importedBookmarks = await importPdfOutlineBookmarks(activePdf, nextPages);
        if (importedBookmarks.length > 0) {
          setBookmarks((existing) => (existing.length > 0 ? existing : importedBookmarks));
        }
        setWorkState(null);
      } catch (error) {
        if (cancelled) return;
        setWorkState(null);
        setErrorMessage(getPdfErrorMessage(error, "render"));
      }
    }

    renderPages();

    return () => {
      cancelled = true;
    };
  }, [pdfDoc, activeSourceDocumentId]);

  async function renderPage(page: PDFPageProxy, sourceDocumentId?: string): Promise<PageView> {
    const viewport = page.getViewport({ scale: PAGE_SCALE });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");

    if (!context) {
      throw new Error("Could not create a canvas context.");
    }

    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport }).promise;
    const textContent = await page.getTextContent();

    return {
      id: crypto.randomUUID(),
      pageNumber: page.pageNumber,
      sourceDocumentId,
      sourcePageNumber: page.pageNumber,
      rotation: 0,
      width: viewport.width,
      height: viewport.height,
      imageUrl: canvas.toDataURL("image/png"),
      textItems: textContent.items.flatMap((item, index) => getTextItemView(item, index, viewport.transform)),
    };
  }

  async function loadFile(file: File) {
    if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
      setErrorMessage("Please choose a PDF file.");
      return;
    }

    try {
      setErrorMessage("");
      setWorkState({ message: `Opening ${file.name}`, progress: 5 });
      const buffer = await file.arrayBuffer();
      const bytes = new Uint8Array(buffer);
      setWorkState({ message: "Checking PDF", progress: 20 });
      const loaded = await getDocument({ data: bytes.slice() }).promise;
      const sourceId = crypto.randomUUID();
      const projectName = file.name.replace(/\.pdf$/i, "") || "Untitled project";

      setWorkState({ message: "Preparing pages", progress: 35 });
      setHasOpenProject(true);
      setProjectId(crypto.randomUUID());
      setProjectMetadata(createProjectMetadata(projectName));
      setSavedProjectFingerprint("");
      setLastSavedAt(null);
      setLastAutosavedAt(null);
      setDesktopProjectPath(null);
      setIsProjectDirty(true);
      setProjectStatusMessage("PDF opened as an untitled project. Use Save Project to create a .pproj file.");
      setPdfName(file.name);
      setPdfBytes(bytes);
      setSourceDocuments({ [sourceId]: { id: sourceId, name: file.name, bytes, mimeType: "application/pdf" } });
      setProjectImageAssets([]);
      setActiveSourceDocumentId(sourceId);
      setPdfDoc(loaded);
      setMarks([]);
      resetNavigationState();
      setPublishingSettings(createDefaultPublishingSettings());
      setPublishingHistory({ past: [], future: [] });
      setHistory({ past: [], future: [] });
      setSelectedMark(null);
      setSelectedPageIds([]);
      setLastSelectedPageId(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(getPdfErrorMessage(error, "upload"));
    }
  }

  function handleUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    if (hasOpenProject && hasUnsavedChanges && !window.confirm("Open this PDF and discard unsaved changes in the current project?")) return;
    void loadFile(file);
  }

  function beginDocxImport(file: File, mode: "replace" | "append") {
    try {
      validateDocxImportFile(file);
      setPendingDocxFile(file);
      setDocxImportMode(mode);
      setDocxImportOptions(defaultDocxImportOptions);
      setDocxReimportTarget(null);
      setErrorMessage("");
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : "The selected Word document could not be imported.");
    }
  }

  function beginDocxReimport(metadata: DocxImportMetadata) {
    const sourceId = metadata.originalSourceDocumentId;
    const source = sourceId ? sourceDocuments[sourceId] : undefined;
    if (!source || source.mimeType !== "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
      setErrorMessage("The original DOCX source is missing from this project. Re-import cannot continue until the source is restored.");
      return;
    }
    const file = new File([source.bytes.slice()], source.name, { type: source.mimeType });
    setPendingDocxFile(file);
    setDocxImportMode("append");
    setDocxImportOptions(normalizeDocxImportOptions(metadata.options ?? defaultDocxImportOptions));
    setDocxReimportTarget(metadata);
    setErrorMessage("");
  }

  function handleDocxUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    const mode = hasOpenProject ? "append" : "replace";
    beginDocxImport(file, mode);
  }

  function beginPptxImport(file: File, mode: "replace" | "append") {
    try {
      validatePptxImportFile(file);
      setPendingPptxFile(file);
      setPptxImportMode(mode);
      setPptxImportOptions(defaultPptxImportOptions);
      setPptxReimportTarget(null);
      setErrorMessage("");
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : "The selected PowerPoint presentation could not be imported.");
    }
  }

  function beginPptxReimport(metadata: PptxImportMetadata) {
    const sourceId = metadata.originalSourceDocumentId;
    const source = sourceId ? sourceDocuments[sourceId] : undefined;
    if (!source || source.mimeType !== "application/vnd.openxmlformats-officedocument.presentationml.presentation") {
      setErrorMessage("The original PPTX source is missing from this project. Re-import cannot continue until the source is restored.");
      return;
    }
    const file = new File([source.bytes.slice()], source.name, { type: source.mimeType });
    setPendingPptxFile(file);
    setPptxImportMode("append");
    setPptxImportOptions(metadata.options ?? defaultPptxImportOptions);
    setPptxReimportTarget(metadata);
    setErrorMessage("");
  }

  function handlePptxUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    beginPptxImport(file, hasOpenProject ? "append" : "replace");
  }

  async function confirmDocxImport() {
    if (!pendingDocxFile) return;
    if (docxImportMode === "replace" && hasOpenProject && hasUnsavedChanges && !window.confirm("Import this Word document and discard unsaved changes in the current project?")) return;
    try {
      setErrorMessage("");
      const result = await importDocxDocument(pendingDocxFile, docxImportOptions, (progress) => setWorkState({ message: progress.stage, progress: progress.progress }));
      setWorkState({ message: "Rendering imported pages", progress: 88 });
      const convertedPdf = await getDocument({ data: result.convertedPdfSource.bytes.slice() }).promise;
      const importedPages: PageView[] = [];
      for (let pageNumber = 1; pageNumber <= convertedPdf.numPages; pageNumber += 1) {
        setWorkState({ message: `Rendering imported page ${pageNumber} of ${convertedPdf.numPages}`, progress: 88 + (pageNumber / convertedPdf.numPages) * 8 });
        const renderedPage = await renderPage(await convertedPdf.getPage(pageNumber), result.convertedPdfSource.id);
        importedPages.push({
          ...renderedPage,
          importMetadata: result.importMetadata,
          sourceMappings: result.sourceMappings.filter((mapping) => mapping.pageNumber === pageNumber),
        });
      }
      const pageIdsByNumber = new Map(importedPages.map((page) => [page.sourcePageNumber ?? page.pageNumber, page.id]));
      const importedBookmarks = docxImportOptions.createBookmarks ? createBookmarksFromDocxHeadings(result.headings, pageIdsByNumber) : [];
      const importMetadata = docxReimportTarget
        ? { ...result.importMetadata, revisionHistory: mergeDocxRevisionHistory(result.importMetadata, docxReimportTarget) }
        : result.importMetadata;
      const importedCommentMarks = createWordCommentMarks(result.wordComments, importedPages);
      const importedNumberingSections = createDocxNumberingSections(result.sectionNumbering, importedPages);
      const importedReport = {
        ...result.report,
        bookmarksCreated: importedBookmarks.length,
        commentsImported: importedCommentMarks.length,
        sectionNumberingDetected: importedNumberingSections.length || result.report.sectionNumberingDetected,
        revision: importMetadata.revisionHistory[0] ?? result.report.revision,
      };
      const imageAssets = result.imageAssets
        .filter((image): image is typeof image & { mimeType: "image/png" | "image/jpeg" } => image.mimeType === "image/png" || image.mimeType === "image/jpeg")
        .map<ProjectImageAsset>((image) => ({
          id: image.id,
          name: image.name,
          dataUrl: image.dataUrl,
          mimeType: image.mimeType,
          naturalWidth: image.width,
          naturalHeight: image.height,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        }));
      const importedSources: Record<string, SourceDocument> = {
        [result.convertedPdfSource.id]: {
          id: result.convertedPdfSource.id,
          name: result.convertedPdfSource.name,
          bytes: result.convertedPdfSource.bytes,
          mimeType: result.convertedPdfSource.mimeType,
          importMetadata,
        },
      };
      if (result.originalSource) {
        importedSources[result.originalSource.id] = {
          id: result.originalSource.id,
          name: result.originalSource.name,
          bytes: result.originalSource.bytes,
          mimeType: result.originalSource.mimeType,
          importMetadata,
        };
      }
      importedPages.forEach((page) => {
        page.importMetadata = importMetadata;
      });

      if (docxReimportTarget && hasOpenProject) {
        const targetPageIndexes = pages
          .map((page, index) => ({ page, index }))
          .filter((item) => item.page.importMetadata?.importId === docxReimportTarget.importId);
        const insertIndex = targetPageIndexes[0]?.index ?? getPageInsertionIndex("after");
        const replacedPageIds = targetPageIndexes.map((item) => item.page.id);
        const oldToNewPageId = new Map(replacedPageIds.map((pageId, index) => [pageId, importedPages[index]?.id]).filter((entry): entry is [string, string] => Boolean(entry[1])));
        const retainedPages = pages.filter((page) => page.importMetadata?.importId !== docxReimportTarget.importId);
        const nextPages = renumberPages([...retainedPages.slice(0, insertIndex), ...importedPages, ...retainedPages.slice(insertIndex)]);
        const nextMarks = marks.map((mark) => {
          const nextPageId = mark.pageId ? oldToNewPageId.get(mark.pageId) : undefined;
          if (!nextPageId) return mark;
          const nextPage = nextPages.find((page) => page.id === nextPageId);
          return nextPage ? { ...mark, pageId: nextPageId, page: nextPage.pageNumber } : mark;
        }).filter((mark) => !mark.pageId || !replacedPageIds.includes(mark.pageId) || oldToNewPageId.has(mark.pageId));
        setSourceDocuments((existing) => ({ ...existing, ...importedSources }));
        setProjectImageAssets((existing) => [...existing, ...imageAssets.filter((asset) => !existing.some((item) => item.dataUrl === asset.dataUrl))]);
        commitDocument(nextPages, [...nextMarks, ...importedCommentMarks], null, insertIndex + 1, importedPages.map((page) => page.id));
        commitNavigation({
          bookmarks: [...bookmarks.filter((bookmark) => !replacedPageIds.includes(bookmark.pageId)), ...importedBookmarks],
          numberingSections: [...numberingSections.filter((section) => !replacedPageIds.includes(section.startPageId)), ...importedNumberingSections],
        });
        setProjectStatusMessage(`Re-imported Word source ${pendingDocxFile.name}. Manual page annotations were preserved where page order still matched.`);
      } else if (docxImportMode === "append" && hasOpenProject) {
        const insertIndex = getPageInsertionIndex("after");
        const nextPages = renumberPages([...pages.slice(0, insertIndex), ...importedPages, ...pages.slice(insertIndex)]);
        setSourceDocuments((existing) => ({ ...existing, ...importedSources }));
        setProjectImageAssets((existing) => [...existing, ...imageAssets.filter((asset) => !existing.some((item) => item.dataUrl === asset.dataUrl))]);
        commitDocument(nextPages, [...marks, ...importedCommentMarks], null, insertIndex + 1, importedPages.map((page) => page.id));
        commitNavigation({ bookmarks: [...bookmarks, ...importedBookmarks], numberingSections: [...numberingSections, ...importedNumberingSections] });
        setProjectStatusMessage(`Imported Word document ${pendingDocxFile.name} into the current project.`);
      } else {
        const projectName = result.projectName;
        setHasOpenProject(true);
        setProjectId(crypto.randomUUID());
        setProjectMetadata(createProjectMetadata(projectName));
        setSavedProjectFingerprint("");
        setLastSavedAt(null);
        setLastAutosavedAt(null);
        setIsProjectDirty(true);
        setProjectStatusMessage("Word document imported as an untitled Publish Pro project. Use Save Project to create a .pproj file.");
        setPdfName(result.defaultPdfName);
        setPdfBytes(result.convertedPdfSource.bytes);
        setPdfDoc(null);
        setActiveSourceDocumentId(result.convertedPdfSource.id);
        setSourceDocuments(importedSources);
        setProjectImageAssets(imageAssets);
        setPages(importedPages);
        setMarks(importedCommentMarks);
        resetNavigationState();
        setBookmarks(importedBookmarks);
        setNumberingSections(importedNumberingSections);
        setPublishingSettings(result.publishingSettingsPatch ?? createDefaultPublishingSettings());
        setPublishingHistory({ past: [], future: [] });
        setHistory({ past: [], future: [] });
        setSelectedMark(null);
        setSelectedPageIds(importedPages[0] ? [importedPages[0].id] : []);
        setLastSelectedPageId(importedPages[0]?.id ?? null);
        setCurrentPage(1);
        setActiveWorkspace("publish");
        setLeftPanel(importedBookmarks.length > 0 ? "bookmarks" : "pages");
      }
      setDocxImportReport(importedReport);
      setPendingDocxFile(null);
      setDocxReimportTarget(null);
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(error instanceof Error ? error.message : "The Word document could not be imported.");
    }
  }

  async function confirmPptxImport() {
    if (!pendingPptxFile) return;
    if (pptxImportMode === "replace" && hasOpenProject && hasUnsavedChanges && !window.confirm("Import this PowerPoint presentation and discard unsaved changes in the current project?")) return;
    try {
      setErrorMessage("");
      const result = await importPptxPresentation(pendingPptxFile, pptxImportOptions, (progress) => setWorkState({ message: progress.stage, progress: progress.progress }));
      setWorkState({ message: "Rendering imported slides", progress: 88 });
      const convertedPdf = await getDocument({ data: result.convertedPdfSource.bytes.slice() }).promise;
      const importedPages: PageView[] = [];
      for (let pageNumber = 1; pageNumber <= convertedPdf.numPages; pageNumber += 1) {
        setWorkState({ message: `Rendering imported slide ${pageNumber} of ${convertedPdf.numPages}`, progress: 88 + (pageNumber / convertedPdf.numPages) * 8 });
        const renderedPage = await renderPage(await convertedPdf.getPage(pageNumber), result.convertedPdfSource.id);
        const slideTitle = result.importMetadata.slideTitles.find((title) => title.pageNumber === pageNumber);
        importedPages.push({
          ...renderedPage,
          label: slideTitle?.hidden ? `${slideTitle.title} (hidden)` : slideTitle?.title,
          importMetadata: result.importMetadata,
          sourceMappings: result.sourceMappings.filter((mapping) => mapping.pageNumber === pageNumber),
        });
      }
      const pageIdsByNumber = new Map(importedPages.map((page) => [page.sourcePageNumber ?? page.pageNumber, page.id]));
      const importedBookmarks = pptxImportOptions.createBookmarks ? result.bookmarks.map((bookmark) => ({ ...bookmark, pageId: pageIdsByNumber.get(bookmark.order) ?? bookmark.pageId })) : [];
      const importMetadata = pptxReimportTarget ? { ...result.importMetadata, revisionHistory: [result.importMetadata.revisionHistory[0], ...(pptxReimportTarget.revisionHistory ?? [])].filter(Boolean).slice(0, 2) } : result.importMetadata;
      importedPages.forEach((page) => {
        page.importMetadata = importMetadata;
      });
      const imageAssets = result.imageAssets
        .filter((image): image is typeof image & { mimeType: "image/png" | "image/jpeg" } => image.mimeType === "image/png" || image.mimeType === "image/jpeg")
        .map<ProjectImageAsset>((image) => ({
          id: image.id,
          name: image.name,
          dataUrl: image.dataUrl,
          mimeType: image.mimeType,
          naturalWidth: image.width,
          naturalHeight: image.height,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        }));
      const importedSources: Record<string, SourceDocument> = {
        [result.convertedPdfSource.id]: { id: result.convertedPdfSource.id, name: result.convertedPdfSource.name, bytes: result.convertedPdfSource.bytes, mimeType: result.convertedPdfSource.mimeType, importMetadata },
      };
      if (result.originalSource) {
        importedSources[result.originalSource.id] = { id: result.originalSource.id, name: result.originalSource.name, bytes: result.originalSource.bytes, mimeType: result.originalSource.mimeType, importMetadata };
      }

      if (pptxReimportTarget && hasOpenProject) {
        const targetPageIndexes = pages.map((page, index) => ({ page, index })).filter((item) => item.page.importMetadata?.importId === pptxReimportTarget.importId);
        const insertIndex = targetPageIndexes[0]?.index ?? getPageInsertionIndex("after");
        const replacedPageIds = targetPageIndexes.map((item) => item.page.id);
        const oldToNewPageId = new Map(replacedPageIds.map((pageId, index) => [pageId, importedPages[index]?.id]).filter((entry): entry is [string, string] => Boolean(entry[1])));
        const retainedPages = pages.filter((page) => page.importMetadata?.importId !== pptxReimportTarget.importId);
        const nextPages = renumberPages([...retainedPages.slice(0, insertIndex), ...importedPages, ...retainedPages.slice(insertIndex)]);
        const nextMarks = marks.map((mark) => {
          const nextPageId = mark.pageId ? oldToNewPageId.get(mark.pageId) : undefined;
          if (!nextPageId) return mark;
          const nextPage = nextPages.find((page) => page.id === nextPageId);
          return nextPage ? { ...mark, pageId: nextPageId, page: nextPage.pageNumber } : mark;
        }).filter((mark) => !mark.pageId || !replacedPageIds.includes(mark.pageId) || oldToNewPageId.has(mark.pageId));
        setSourceDocuments((existing) => ({ ...existing, ...importedSources }));
        setProjectImageAssets((existing) => [...existing, ...imageAssets.filter((asset) => !existing.some((item) => item.dataUrl === asset.dataUrl))]);
        commitDocument(nextPages, nextMarks, null, insertIndex + 1, importedPages.map((page) => page.id));
        commitNavigation({ bookmarks: [...bookmarks.filter((bookmark) => !replacedPageIds.includes(bookmark.pageId)), ...importedBookmarks] });
        setProjectStatusMessage(`Re-imported PowerPoint source ${pendingPptxFile.name}. Manual slide annotations were preserved where slide order still matched.`);
      } else if (pptxImportMode === "append" && hasOpenProject) {
        const insertIndex = getPageInsertionIndex("after");
        const nextPages = renumberPages([...pages.slice(0, insertIndex), ...importedPages, ...pages.slice(insertIndex)]);
        setSourceDocuments((existing) => ({ ...existing, ...importedSources }));
        setProjectImageAssets((existing) => [...existing, ...imageAssets.filter((asset) => !existing.some((item) => item.dataUrl === asset.dataUrl))]);
        commitDocument(nextPages, marks, null, insertIndex + 1, importedPages.map((page) => page.id));
        commitNavigation({ bookmarks: [...bookmarks, ...importedBookmarks] });
        setProjectStatusMessage(`Imported PowerPoint presentation ${pendingPptxFile.name} into the current project.`);
      } else {
        setHasOpenProject(true);
        setProjectId(crypto.randomUUID());
        setProjectMetadata(createProjectMetadata(result.projectName));
        setSavedProjectFingerprint("");
        setLastSavedAt(null);
        setLastAutosavedAt(null);
        setIsProjectDirty(true);
        setProjectStatusMessage("PowerPoint presentation imported as an untitled Publish Pro project. Use Save Project to create a .pproj file.");
        setPdfName(result.defaultPdfName);
        setPdfBytes(result.convertedPdfSource.bytes);
        setPdfDoc(null);
        setActiveSourceDocumentId(result.convertedPdfSource.id);
        setSourceDocuments(importedSources);
        setProjectImageAssets(imageAssets);
        setPages(importedPages);
        setMarks([]);
        resetNavigationState();
        setBookmarks(importedBookmarks);
        setPublishingSettings(createDefaultPublishingSettings());
        setPublishingHistory({ past: [], future: [] });
        setHistory({ past: [], future: [] });
        setSelectedMark(null);
        setSelectedPageIds(importedPages[0] ? [importedPages[0].id] : []);
        setLastSelectedPageId(importedPages[0]?.id ?? null);
        setCurrentPage(1);
        setActiveWorkspace("assemble");
        setLeftPanel(importedBookmarks.length > 0 ? "bookmarks" : "pages");
      }
      setPptxImportReport({ ...result.report, slidesImported: importedPages.length });
      setPendingPptxFile(null);
      setPptxReimportTarget(null);
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(error instanceof Error ? error.message : "The PowerPoint presentation could not be imported.");
    }
  }

  async function handleImageUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    try {
      setErrorMessage("");
      setWorkState({ message: `Preparing ${file.name}`, progress: 15 });
      const image = await prepareImageAsset(file);
      const page = pages.find((item) => item.pageNumber === currentPage) ?? pages[0];
      const maxWidth = page.width * 0.36;
      const maxHeight = page.height * 0.28;
      const scale = Math.min(maxWidth / image.width, maxHeight / image.height, 1);
      const width = Math.max(48, image.width * scale);
      const height = Math.max(48, image.height * scale);
      const mark: Mark = {
        id: crypto.randomUUID(),
        kind: "image",
        pageId: getPageIdForNumber(currentPage),
        page: currentPage,
        x: (page.width - width) / 2,
        y: (page.height - height) / 2,
        width,
        height,
        text: image.name,
        color: "#111827",
        size: 16,
        rotation: 0,
        imageDataUrl: image.dataUrl,
        imageMimeType: image.mimeType,
        imageName: image.name,
        imageNaturalWidth: image.width,
        imageNaturalHeight: image.height,
      };

      commitMarks([...marks, clampMarkToPage(mark, pages)], mark.id);
      setActiveTool("select");
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(getImageErrorMessage(error));
    }
  }

  async function importImageAsset(file: File) {
    const image = await prepareImageAsset(file);
    const existing = projectImageAssets.find((asset) => asset.dataUrl === image.dataUrl);
    if (existing) return existing;
    const existingMarkAsset = marks.find((mark) => (mark.kind === "image" || mark.kind === "pngSignature") && mark.imageDataUrl === image.dataUrl);
    if (existingMarkAsset) {
      return {
        id: existingMarkAsset.id,
        name: existingMarkAsset.imageName || existingMarkAsset.text || image.name,
        dataUrl: existingMarkAsset.imageDataUrl ?? image.dataUrl,
        mimeType: existingMarkAsset.imageMimeType ?? image.mimeType,
        naturalWidth: existingMarkAsset.imageNaturalWidth ?? image.width,
        naturalHeight: existingMarkAsset.imageNaturalHeight ?? image.height,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
    }

    const now = new Date().toISOString();
    const asset: ProjectImageAsset = {
      id: crypto.randomUUID(),
      name: image.name,
      dataUrl: image.dataUrl,
      mimeType: image.mimeType,
      naturalWidth: image.width,
      naturalHeight: image.height,
      createdAt: now,
      updatedAt: now,
    };
    setProjectImageAssets((existingAssets) => {
      if (existingAssets.some((item) => item.dataUrl === image.dataUrl)) return existingAssets;
      return [...existingAssets, asset];
    });
    setIsProjectDirty(true);
    return asset;
  }

  function startProjectImageAssetImport(target: ProjectAssetImportTarget) {
    projectAssetImportTarget.current = target;
    projectImageInput.current?.click();
  }

  async function handleProjectImageAssetUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    const target = projectAssetImportTarget.current;
    projectAssetImportTarget.current = null;

    try {
      setErrorMessage("");
      setWorkState({ message: `Importing ${file.name}`, progress: 20 });
      const asset = await importImageAsset(file);
      if (target?.kind === "watermark") {
        updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, imageAssetId: asset.id } }));
      }
      if (target?.kind === "headerFooter") {
        updateHeaderFooterZoneImage(target.area, target.side, { assetId: asset.id });
      }
      if (target?.kind === "headerFooterOverride") {
        updateHeaderFooterOverrideImage(target.key, target.zone, { assetId: asset.id });
      }
      setProjectStatusMessage(`${asset.name} imported to Project Assets.`);
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(getImageErrorMessage(error));
    }
  }

  async function handleSignatureUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    try {
      setErrorMessage("");
      setWorkState({ message: `Preparing signature ${file.name}`, progress: 20 });
      const signature = await preparePngSignature(file);
      setPendingSignature(signature);
      setActiveTool("pngSignature");
      setSelectedMark(null);
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setPendingSignature(null);
      setActiveTool("select");
      setErrorMessage(getSignatureErrorMessage(error));
    }
  }

  function addMark(pageNumber: number, x: number, y: number) {
    if (activeTool === "select") return;
    if (activeTool === "text") return;
    if (activeTool === "image") {
      imageInput.current?.click();
      return;
    }
    if (activeTool === "pngSignature") {
      signatureInput.current?.click();
      return;
    }
    if (activeTool === "draw") return;
    if (isShapeTool(activeTool)) return;
    if (isTextMarkupKind(activeTool)) return;

    if (activeTool === "comment") {
      const comment = createCommentData();
      const mark: Mark = {
        id: crypto.randomUUID(),
        kind: "comment",
        pageId: getPageIdForNumber(pageNumber),
        page: pageNumber,
        x,
        y,
        width: 24,
        height: 24,
        text: comment.title ?? "Comment",
        color: comment.color,
        size: 16,
        rotation: 0,
        comment,
      };

      commitMarks([...marks, clampMarkToPage(mark, pages)], mark.id);
      setActiveTool("select");
      return;
    }

    const mark: Mark = {
      id: crypto.randomUUID(),
      kind: activeTool,
      pageId: getPageIdForNumber(pageNumber),
      page: pageNumber,
      x,
      y,
      width: activeTool === "highlight" ? 220 : activeTool === "signature" ? 190 : 124,
      height: activeTool === "highlight" ? 28 : 42,
      text:
        activeTool === "signature"
            ? "Signature"
            : activeTool === "stamp"
              ? "APPROVED"
              : "",
      color: activeTool === "highlight" ? "#facc15" : activeTool === "stamp" ? BRAND_RED : "#111827",
      size: 16,
      rotation: 0,
    };

    commitMarks([...marks, mark], mark.id);
    setActiveTool("select");
  }

  function addShape(pageNumber: number, placement: { x: number; y: number; width?: number; height?: number }, shapeKind: ShapeKind) {
    const isLineShape = shapeKind === "line" || shapeKind === "arrow" || shapeKind === "doubleArrow";
    const width = Math.max(isLineShape ? 64 : 48, placement.width ?? (isLineShape ? 160 : 150));
    const height = Math.max(isLineShape ? 8 : 36, placement.height ?? (isLineShape ? 44 : 96));
    const mark: Mark = {
      id: crypto.randomUUID(),
      kind: "shape",
      pageId: getPageIdForNumber(pageNumber),
      page: pageNumber,
      x: placement.x,
      y: placement.y,
      width,
      height,
      text: shapeToolLabels[shapeKind],
      color: "#d8342a",
      size: 16,
      rotation: 0,
      opacity: 1,
      shapeStyle: createShapeStyle(shapeKind),
    };

    commitMarks([...marks, clampMarkToPage(mark, pages)], mark.id);
    setActiveTool("select");
  }

  function addTextBox(pageNumber: number, x: number, y: number, width?: number, height?: number) {
    if (activeTool !== "text") return;
    const page = pages.find((item) => item.pageNumber === pageNumber) ?? pages[0];
    const boxWidth = Math.max(96, width ?? 190);
    const boxHeight = Math.max(42, height ?? 64);
    const mark: Mark = {
      id: crypto.randomUUID(),
      kind: "text",
      pageId: getPageIdForNumber(pageNumber),
      page: pageNumber,
      x: width === undefined ? x : Math.min(x, page.width - boxWidth),
      y: height === undefined ? y : Math.min(y, page.height - boxHeight),
      width: boxWidth,
      height: boxHeight,
      text: "",
      color: "#111827",
      size: 12,
      rotation: 0,
      opacity: 1,
      textStyle: createDefaultTextBoxStyle(),
    };
    const clamped = clampMarkToPage(mark, pages);

    commitMarks([...marks, clamped], clamped.id);
    setEditingText({ id: clamped.id, before: clamped, draft: "", isNew: true });
    setActiveTool("select");
  }

  function addTextMarkup(pageNumber: number, rects: MarkupRect[]) {
    if (!isTextMarkupKind(activeTool)) return;
    const bounds = getMarkupBounds(rects);
    if (!bounds) return;

    const mark: Mark = {
      id: crypto.randomUUID(),
      kind: activeTool,
      pageId: getPageIdForNumber(pageNumber),
      page: pageNumber,
      x: bounds.x,
      y: bounds.y,
      width: bounds.width,
      height: bounds.height,
      text: getMarkupToolLabel(activeTool),
      color: activeTool === "textHighlight" ? "#facc15" : "#d8342a",
      size: 16,
      rotation: 0,
      opacity: activeTool === "textHighlight" ? 0.38 : 1,
      thickness: activeTool === "textHighlight" ? undefined : 2,
      markupRects: rects,
    };

    commitMarks([...marks, clampMarkToPage(mark, pages)], mark.id);
  }

  function placeSignature(pageNumber: number, x: number, y: number, width?: number, height?: number) {
    if (!pendingSignature) return;
    const page = pages.find((item) => item.pageNumber === pageNumber) ?? pages[0];
    const defaultSize = getDefaultImageMarkSize(pendingSignature, page, 0.32, 0.18);
    const aspectRatio = pendingSignature.width / pendingSignature.height || 1;
    const markWidth = Math.max(24, width ?? defaultSize.width);
    const markHeight = Math.max(12, height ?? markWidth / aspectRatio);
    const mark: Mark = {
      id: crypto.randomUUID(),
      kind: "pngSignature",
      pageId: getPageIdForNumber(pageNumber),
      page: pageNumber,
      x: width === undefined ? x - markWidth / 2 : x,
      y: height === undefined ? y - markHeight / 2 : y,
      width: markWidth,
      height: markHeight,
      text: pendingSignature.name,
      color: "#111827",
      size: 16,
      rotation: 0,
      imageDataUrl: pendingSignature.dataUrl,
      imageMimeType: "image/png",
      imageName: pendingSignature.name,
      imageNaturalWidth: pendingSignature.width,
      imageNaturalHeight: pendingSignature.height,
      opacity: 1,
      lockAspectRatio: true,
    };

    commitMarks([...marks, clampMarkToPage(mark, pages)], mark.id);
    setPendingSignature(null);
    setActiveTool("select");
  }

  function updateMark(id: string, patch: Partial<Mark>) {
    commitMarks(marks.map((mark) => (mark.id === id ? clampMarkToPage({ ...mark, ...patch }, pages) : mark)), selectedMark);
  }

  function previewMark(id: string, patch: Partial<Mark>) {
    setMarks((existing) => existing.map((mark) => (mark.id === id ? clampMarkToPage({ ...mark, ...patch }, pages) : mark)));
  }

  function commitMarkChange(beforeMark: Mark, afterMark: Mark) {
    const nextMarks = marks.map((mark) => (mark.id === afterMark.id ? clampMarkToPage(afterMark, pages) : mark));
    const beforeMarks = marks.map((mark) => (mark.id === beforeMark.id ? beforeMark : mark));
    commitMarks(nextMarks, selectedMark, beforeMarks);
  }

  function addStroke(pageNumber: number, points: StrokePoint[]) {
    const smoothedPoints = smoothStrokePoints(points);
    const bounds = getStrokeBounds(smoothedPoints, penWidth);
    if (!bounds) return;

    const stroke: Mark = {
      id: crypto.randomUUID(),
      kind: "stroke",
      pageId: getPageIdForNumber(pageNumber),
      page: pageNumber,
      x: bounds.x,
      y: bounds.y,
      width: bounds.width,
      height: bounds.height,
      text: "Freehand stroke",
      color: penColor,
      size: penWidth,
      rotation: 0,
      strokePoints: smoothedPoints,
      strokeOpacity: penOpacity,
    };

    commitMarks([...marks, clampMarkToPage(stroke, pages)], stroke.id);
  }

  function removeSelectedMark() {
    if (!selectedMark) return;
    commitMarks(marks.filter((mark) => mark.id !== selectedMark), null);
  }

  function removeMarkById(id: string) {
    commitMarks(marks.filter((mark) => mark.id !== id), null);
  }

  function duplicateSelectedMark() {
    if (!selected) return;
    const duplicate = duplicateMark(selected, currentPage);
    commitMarks([...marks, clampMarkToPage(withPageId(duplicate, currentPage), pages)], duplicate.id);
  }

  function duplicateMarkById(id: string) {
    const mark = marks.find((item) => item.id === id);
    if (!mark) return;
    const duplicate = duplicateMark(mark, currentPage);
    commitMarks([...marks, clampMarkToPage(withPageId(duplicate, currentPage), pages)], duplicate.id);
  }

  function pasteCopiedMark() {
    if (!copiedMark) return;
    const pasted = duplicateMark(copiedMark, currentPage);
    commitMarks([...marks, clampMarkToPage(withPageId(pasted, currentPage), pages)], pasted.id);
  }

  function startTextEdit(mark: Mark) {
    if (mark.kind !== "text") return;
    setSelectedMark(mark.id);
    setEditingText({ id: mark.id, before: mark, draft: mark.text, isNew: false });
  }

  function saveTextEdit() {
    if (!editingText) return;
    const current = marks.find((mark) => mark.id === editingText.id);
    if (!current || current.kind !== "text") {
      setEditingText(null);
      return;
    }

    const draft = editingText.draft;
    if (!draft.trim()) {
      commitMarks(marks.filter((mark) => mark.id !== editingText.id), null);
      setEditingText(null);
      return;
    }

    const nextMark = { ...current, text: draft };
    if (areMarksEqual([editingText.before], [nextMark])) {
      setEditingText(null);
      return;
    }

    commitMarkChange(editingText.before, nextMark);
    setEditingText(null);
  }

  function cancelTextEdit() {
    if (!editingText) return;
    if (editingText.isNew) {
      commitMarks(marks.filter((mark) => mark.id !== editingText.id), null);
    } else {
      const nextMarks = marks.map((mark) => (mark.id === editingText.id ? editingText.before : mark));
      commitMarks(nextMarks, editingText.before.id);
    }
    setEditingText(null);
  }

  function updateTextStyle(id: string, patch: Partial<TextBoxStyle>) {
    const mark = marks.find((item) => item.id === id);
    if (!mark) return;
    updateMark(id, { textStyle: { ...createDefaultTextBoxStyle(mark.textStyle), ...patch } });
  }

  function applyTextPreset(preset: string) {
    if (!selected || selected.kind !== "text") return;
    const presetPatch = getPresetTextStyle(preset as TextBoxPreset);
    const { text, color, size, ...stylePatch } = presetPatch;
    updateMark(selected.id, {
      text: text ?? selected.text,
      color: color ?? selected.color,
      size: size ?? selected.size,
      textStyle: { ...createDefaultTextBoxStyle(selected.textStyle), ...stylePatch },
    });
  }

  function updateShapeStyle(id: string, patch: Partial<ShapeStyle>) {
    const mark = marks.find((item) => item.id === id);
    if (!mark?.shapeStyle) return;
    updateMark(id, {
      color: patch.strokeColor ?? mark.color,
      shapeStyle: {
        ...mark.shapeStyle,
        ...patch,
      },
    });
  }

  function saveCommentBody() {
    if (!selected || selected.kind !== "comment" || !selected.comment) return;
    updateMark(selected.id, {
      text: selected.comment.title || "Comment",
      comment: {
        ...selected.comment,
        body: commentDraft.trim(),
        updatedAt: new Date().toISOString(),
      },
    });
  }

  function cancelCommentEdit() {
    if (selected?.kind === "comment" && selected.comment) {
      setCommentDraft(selected.comment.body);
    }
  }

  function updateSelectedComment(patch: Partial<CommentData>) {
    if (!selected || selected.kind !== "comment" || !selected.comment) return;
    const nextComment = {
      ...selected.comment,
      ...patch,
      updatedAt: new Date().toISOString(),
    };
    updateMark(selected.id, {
      color: nextComment.color,
      comment: nextComment,
    });
  }

  function addCommentReply() {
    if (!selected || selected.kind !== "comment" || !selected.comment) return;
    const body = commentReplyDraft.trim();
    if (!body) return;
    updateSelectedComment({
      replies: [...selected.comment.replies, createCommentReply(body)],
    });
    setCommentReplyDraft("");
  }

  function saveReplyEdit(reply: CommentReply) {
    if (!selected || selected.kind !== "comment" || !selected.comment || !editingReply) return;
    const nextReplies = selected.comment.replies.map((item) =>
      item.id === reply.id ? { ...item, body: editingReply.body.trim(), updatedAt: new Date().toISOString() } : item
    );
    updateSelectedComment({ replies: nextReplies });
    setEditingReply(null);
  }

  function deleteReply(replyId: string) {
    if (!selected || selected.kind !== "comment" || !selected.comment) return;
    updateSelectedComment({
      replies: selected.comment.replies.filter((reply) => reply.id !== replyId),
    });
    if (editingReply?.id === replyId) setEditingReply(null);
  }

  function selectComment(mark: Mark) {
    const page = pages.find((item) => item.id === markPageId(mark)) ?? pages.find((item) => item.pageNumber === mark.page);
    setCurrentPage(page?.pageNumber ?? mark.page);
    setSelectedMark(mark.id);
  }

  function handlePageSelect(page: PageView, event: ReactMouseEvent<HTMLButtonElement>) {
    setCurrentPage(page.pageNumber);
    setSelectedMark(null);
    if (event.shiftKey && lastSelectedPageId) {
      const start = pages.findIndex((item) => item.id === lastSelectedPageId);
      const end = pages.findIndex((item) => item.id === page.id);
      if (start >= 0 && end >= 0) {
        const [from, to] = start < end ? [start, end] : [end, start];
        setSelectedPageIds(pages.slice(from, to + 1).map((item) => item.id));
        return;
      }
    }
    if (event.metaKey || event.ctrlKey) {
      setSelectedPageIds((existing) => {
        const next = existing.includes(page.id) ? existing.filter((id) => id !== page.id) : [...existing, page.id];
        return next.length > 0 ? next : [page.id];
      });
      setLastSelectedPageId(page.id);
      return;
    }
    setSelectedPageIds([page.id]);
    setLastSelectedPageId(page.id);
  }

  function openPageContextMenu(page: PageView, event: ReactMouseEvent<HTMLButtonElement>) {
    event.preventDefault();
    if (!selectedPageIds.includes(page.id)) {
      setSelectedPageIds([page.id]);
      setLastSelectedPageId(page.id);
      setCurrentPage(page.pageNumber);
    }
    setPageContextMenu({
      x: Math.min(event.clientX, window.innerWidth - 230),
      y: Math.min(event.clientY, window.innerHeight - 360),
      pageId: page.id,
    });
  }

  function runPageMenuAction(action: () => void | Promise<void>) {
    setPageContextMenu(null);
    void action();
  }

  function handlePageThumbKeyDown(page: PageView, event: ReactKeyboardEvent<HTMLButtonElement>) {
    const index = pages.findIndex((item) => item.id === page.id);
    if (event.key === "ArrowDown" || event.key === "ArrowRight") {
      event.preventDefault();
      const next = pages[Math.min(pages.length - 1, index + 1)];
      if (next) {
        setCurrentPage(next.pageNumber);
        setSelectedPageIds([next.id]);
        setLastSelectedPageId(next.id);
      }
    }
    if (event.key === "ArrowUp" || event.key === "ArrowLeft") {
      event.preventDefault();
      const previous = pages[Math.max(0, index - 1)];
      if (previous) {
        setCurrentPage(previous.pageNumber);
        setSelectedPageIds([previous.id]);
        setLastSelectedPageId(previous.id);
      }
    }
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      handlePageSelect(page, event as unknown as ReactMouseEvent<HTMLButtonElement>);
    }
  }

  function selectAllPages() {
    setSelectedPageIds(pages.map((page) => page.id));
    setLastSelectedPageId(pages[0]?.id ?? null);
  }

  function clearPageSelection() {
    const current = pages.find((page) => page.pageNumber === currentPage) ?? pages[0];
    setSelectedPageIds(current ? [current.id] : []);
    setLastSelectedPageId(current?.id ?? null);
  }

  function getActivePageIds() {
    if (selectedPageIds.length > 0) return selectedPageIds.filter((id) => pages.some((page) => page.id === id));
    const current = pages.find((page) => page.pageNumber === currentPage) ?? pages[0];
    return current ? [current.id] : [];
  }

  function addPage() {
    insertBlankPages("end");
  }

  function addBookmarkForPage(page: PageView, parentId?: string) {
    const bookmark = createBookmark(pageToReference(page), undefined, parentId, bookmarks.filter((item) => item.parentId === parentId).length);
    commitNavigation({ bookmarks: [...bookmarks, bookmark] });
    setSelectedBookmarkId(bookmark.id);
    setLeftPanel("bookmarks");
  }

  function addBookmarkFromCurrentPage() {
    const page = pages.find((item) => item.pageNumber === currentPage) ?? pages[0];
    if (page) addBookmarkForPage(page);
  }

  function addChildBookmark() {
    const parent = bookmarks.find((bookmark) => bookmark.id === selectedBookmarkId);
    const page = parent ? pages.find((item) => item.id === parent.pageId) : pages.find((item) => item.pageNumber === currentPage) ?? pages[0];
    if (page) addBookmarkForPage(page, parent?.id);
  }

  function addBookmarksFromSelectedPages() {
    const activeIds = getActivePageIds();
    const selectedPages = pages.filter((page) => activeIds.includes(page.id));
    const nextBookmarks = [
      ...bookmarks,
      ...selectedPages.map((page, index) => createBookmark(pageToReference(page), page.label || `Page ${page.pageNumber}`, undefined, bookmarks.filter((item) => !item.parentId).length + index)),
    ];
    commitNavigation({ bookmarks: nextBookmarks });
    setLeftPanel("bookmarks");
  }

  function addBookmarksFromPageLabels() {
    const labelledPages = pages.filter((page) => page.label?.trim());
    const nextBookmarks = [
      ...bookmarks,
      ...labelledPages.map((page, index) => createBookmark(pageToReference(page), page.label, undefined, bookmarks.filter((item) => !item.parentId).length + index)),
    ];
    commitNavigation({ bookmarks: nextBookmarks });
    setLeftPanel("bookmarks");
  }

  function updateBookmark(bookmarkId: string, patch: Partial<DocumentBookmark>) {
    commitNavigation({
      bookmarks: bookmarks.map((bookmark) => (bookmark.id === bookmarkId ? { ...bookmark, ...patch, updatedAt: new Date().toISOString() } : bookmark)),
    });
  }

  function removeBookmark(bookmarkId: string) {
    const descendants = getBookmarkDescendants(bookmarks, bookmarkId);
    const mode = descendants.length > 0 && window.confirm("Promote child bookmarks instead of deleting them? Choose Cancel to delete the whole branch.") ? "promoteChildren" : "deleteChildren";
    commitNavigation({ bookmarks: deleteBookmark(bookmarks, bookmarkId, mode) });
    setSelectedBookmarkId(null);
  }

  function repairBookmarkDestination(bookmarkId: string) {
    const page = pages.find((item) => item.pageNumber === currentPage) ?? pages[0];
    if (!page) return;
    updateBookmark(bookmarkId, { pageId: page.id, y: 0 });
  }

  function duplicateBookmark(bookmarkId: string) {
    commitNavigation({ bookmarks: duplicateBookmarkTree(bookmarks, bookmarkId) });
  }

  function nestSelectedBookmark() {
    if (!selectedBookmarkId) return;
    commitNavigation({ bookmarks: nestBookmark(bookmarks, selectedBookmarkId) });
  }

  function outdentSelectedBookmark() {
    if (!selectedBookmarkId) return;
    commitNavigation({ bookmarks: outdentBookmark(bookmarks, selectedBookmarkId) });
  }

  function moveSelectedBookmark(direction: -1 | 1) {
    if (!selectedBookmarkId) return;
    commitNavigation({ bookmarks: moveBookmark(bookmarks, selectedBookmarkId, direction) });
  }

  function navigateToBookmark(bookmark: DocumentBookmark) {
    const page = pages.find((item) => item.id === bookmark.pageId);
    if (!page) return;
    setCurrentPage(page.pageNumber);
    setSelectedPageIds([page.id]);
    setLastSelectedPageId(page.id);
    setSelectedBookmarkId(bookmark.id);
  }

  function duplicatePage() {
    duplicateSelectedPages();
  }

  function duplicateSelectedPages() {
    const activeIds = getActivePageIds();
    if (activeIds.length === 0) return;
    const sourcePages = pages.filter((page) => activeIds.includes(page.id));
    const insertIndex = Math.max(...sourcePages.map((page) => pages.findIndex((item) => item.id === page.id))) + 1;
    const pageIdMap = new Map<string, string>();
    const pageCopies = sourcePages.map((page) => {
      const nextId = crypto.randomUUID();
      pageIdMap.set(page.id, nextId);
      return { ...page, id: nextId, label: page.label ? `${page.label} Copy` : page.label };
    });
    const markCopies = marksForPageIds(activeIds)
      .map((mark): Mark => {
        const oldPageId = markPageId(mark);
        const nextPageId = oldPageId ? pageIdMap.get(oldPageId) : undefined;
        const nextPageNumber = nextPageId ? insertIndex + pageCopies.findIndex((page) => page.id === nextPageId) + 1 : mark.page;
        return {
          ...duplicateMark(mark, nextPageNumber),
          pageId: nextPageId,
          comment: mark.comment ? cloneComment(mark.comment) : mark.comment,
        };
      })
      .filter((mark): mark is Mark => Boolean(mark.pageId));
    const nextPages = renumberPages([...pages.slice(0, insertIndex), ...pageCopies, ...pages.slice(insertIndex)]);
    const nextMarks = [...marks, ...markCopies];
    const nextSelectedIds = pageCopies.map((page) => page.id);
    commitDocument(nextPages, nextMarks, nextSelectedIds.length ? markCopies[0]?.id ?? null : selectedMark, insertIndex + 1, nextSelectedIds);
  }

  function deleteSelectedPages() {
    const activeIds = getActivePageIds();
    if (activeIds.length === 0) return;
    if (activeIds.length > 1 && !window.confirm(`Delete ${activeIds.length} selected pages?`)) return;
    let remainingPages = pages.filter((page) => !activeIds.includes(page.id));
    if (remainingPages.length === 0) remainingPages = [createBlankPageView(1, getBlankPageSize("letterPortrait"), "Blank")];
    const nextPages = renumberPages(remainingPages);
    const deletedIds = new Set(activeIds);
    const nextMarks = marks.filter((mark) => {
      const id = markPageId(mark);
      return id ? !deletedIds.has(id) : true;
    });
    const nextPage = Math.min(currentPage, nextPages.length);
    const selectedPage = nextPages[nextPage - 1] ?? nextPages[0];
    commitDocument(nextPages, nextMarks, null, selectedPage.pageNumber, [selectedPage.id]);
  }

  async function rotateSelectedPages(delta: 90 | -90 | 180) {
    const activeIds = getActivePageIds();
    if (activeIds.length === 0) return;
    setWorkState({ message: "Rotating pages", progress: 20 });
    const rotatedPages = await Promise.all(pages.map((page) => (activeIds.includes(page.id) ? rotatePageView(page, delta) : page)));
    const rotatedIds = new Set(activeIds);
    const nextMarks = marks.map((mark) => {
      const pageId = markPageId(mark);
      const page = pageId ? pages.find((item) => item.id === pageId) : undefined;
      return page && rotatedIds.has(page.id) ? rotateMarkForPage({ ...mark, pageId: page.id, page: page.pageNumber }, page, delta) : mark;
    });
    setWorkState(null);
    commitDocument(rotatedPages, nextMarks, selectedMark, currentPage, activeIds);
  }

  function insertBlankPages(position: "before" | "after" | "start" | "end") {
    const targetIndex = getPageInsertionIndex(position);
    const current = pages.find((page) => page.pageNumber === currentPage) ?? pages[0];
    const size =
      blankPagePreset === "matchCurrent" && current
        ? { width: current.width / PAGE_SCALE, height: current.height / PAGE_SCALE }
        : blankPagePreset === "custom"
        ? { width: dimensionToPoints(customPageWidth, blankPageUnit), height: dimensionToPoints(customPageHeight, blankPageUnit) }
        : getBlankPageSize(blankPagePreset);
    if (size.width < 72 || size.height < 72 || size.width > 2880 || size.height > 2880) {
      setErrorMessage("Blank page dimensions must be between 72 and 2880 points.");
      return;
    }
    const quantity = clamp(Math.round(blankPageQuantity), 1, 50);
    const orientedSize =
      blankPageOrientation === "landscape" && size.height > size.width
        ? { width: size.height, height: size.width }
        : blankPageOrientation === "portrait" && size.width > size.height
        ? { width: size.height, height: size.width }
        : size;
    const blanks = Array.from({ length: quantity }, (_, index) => ({
      ...createBlankPageView(0, orientedSize, blankPageLabel ? `${blankPageLabel}${quantity > 1 ? ` ${index + 1}` : ""}` : undefined),
      background: blankPageBackground || "#ffffff",
    }));
    const nextPages = renumberPages([...pages.slice(0, targetIndex), ...blanks, ...pages.slice(targetIndex)]);
    const selectedIds = blanks.map((page) => page.id);
    commitDocument(nextPages, marks, null, targetIndex + 1, selectedIds);
  }

  function insertOrRegenerateToc() {
    const visiblePages = pageReferences.filter((page) => !generatedToc?.pageIds.includes(page.id));
    const layout = layoutToc(tocEntries, visiblePages, tocSettings, getResolvedPageTextMap(visiblePages));
    const existingTocIds = new Set(generatedToc?.pageIds ?? []);
    const pagesWithoutOldToc = pages.filter((page) => !existingTocIds.has(page.id));
    const insertIndex = getTocInsertionIndex(pagesWithoutOldToc);
    const tocPageSize = getTocPageSize(tocSettings.pageSize);
    const tocPages: PageView[] = layout.pages.map((tocPage, index) => ({
      ...createBlankPageView(0, tocPageSize, index === 0 ? tocSettings.title : `${tocSettings.title} ${index + 1}`),
      id: tocPage.id,
      generatedToc: true,
      imageUrl: createTocPreviewDataUrl(layout, tocPage.id, tocSettings),
    }));
    const nextPages = renumberPages([...pagesWithoutOldToc.slice(0, insertIndex), ...tocPages, ...pagesWithoutOldToc.slice(insertIndex)]);
    const nextGeneratedToc: TocGeneratedPageMetadata = {
      pageIds: tocPages.map((page) => page.id),
      insertedAtPageId: nextPages[Math.min(insertIndex + tocPages.length, nextPages.length - 1)]?.id,
      generatedAt: new Date().toISOString(),
      stale: false,
    };
    commitDocument(nextPages, marks, selectedMark, insertIndex + 1, tocPages.map((page) => page.id), pages, marks);
    commitNavigation({ generatedToc: nextGeneratedToc }, false);
  }

  function removeGeneratedToc() {
    if (!generatedToc?.pageIds.length) return;
    const tocIds = new Set(generatedToc.pageIds);
    const remainingPages = pages.filter((page) => !tocIds.has(page.id));
    const nextPages = renumberPages(remainingPages.length > 0 ? remainingPages : [createBlankPageView(1, getBlankPageSize("letterPortrait"), "Blank")]);
    commitDocument(nextPages, marks, null, Math.min(currentPage, nextPages.length), nextPages[0] ? [nextPages[0].id] : []);
    commitNavigation({ generatedToc: undefined }, false);
  }

  function addManualTocEntry() {
    const page = pages.find((item) => item.pageNumber === currentPage) ?? pages[0];
    if (!page) return;
    const now = new Date().toISOString();
    const entry: TocManualEntry = {
      id: crypto.randomUUID(),
      title: page.label || `Page ${page.pageNumber}`,
      pageId: page.id,
      level: 1,
      order: manualTocEntries.length,
      enabled: true,
      createdAt: now,
      updatedAt: now,
    };
    commitNavigation({ manualTocEntries: [...manualTocEntries, entry] });
  }

  function addNumberingSection(label = "Section") {
    const page = pages.find((item) => item.pageNumber === currentPage) ?? pages[0];
    if (!page) return;
    commitNavigation({ numberingSections: [...numberingSections, createNumberingSection(page.id, label)] });
  }

  function getResolvedPageTextMap(pageSet: PageReference[]) {
    return new Map(pageSet.map((page) => [page.id, getResolvedPageText(page, resolvedPageNumbers)]));
  }

  function updateNumberingSection(sectionId: string, patch: Partial<NumberingSection>) {
    commitNavigation({
      numberingSections: numberingSections.map((section) => (section.id === sectionId ? { ...section, ...patch } : section)),
    });
  }

  function deleteNumberingSection(sectionId: string) {
    commitNavigation({ numberingSections: numberingSections.filter((section) => section.id !== sectionId) });
  }

  function duplicateSection(section: NumberingSection) {
    commitNavigation({ numberingSections: [...numberingSections, duplicateNumberingSection(section)] });
  }

  function moveNumberingSection(sectionId: string, direction: -1 | 1) {
    const index = numberingSections.findIndex((section) => section.id === sectionId);
    const target = index + direction;
    if (index < 0 || target < 0 || target >= numberingSections.length) return;
    const next = [...numberingSections];
    const [section] = next.splice(index, 1);
    next.splice(target, 0, section);
    commitNavigation({ numberingSections: next });
  }

  function createDefaultNumberingWorkflow() {
    if (pages.length === 0) return;
    const tocPages = pages.filter((page) => page.generatedToc);
    const firstNonToc = pages.find((page) => !page.generatedToc && page.pageNumber > 1) ?? pages[1] ?? pages[0];
    const appendixStart = pages.find((page) => /appendix/i.test(page.label ?? "")) ?? pages[pages.length - 3] ?? firstNonToc;
    const sections: NumberingSection[] = [
      { ...createNumberingSection(pages[0].id, "Cover"), includeNumbering: false, includeInTotal: false, endPageId: pages[0].id },
    ];
    if (tocPages.length > 0) {
      sections.push({
        ...createNumberingSection(tocPages[0].id, "TOC"),
        endPageId: tocPages[tocPages.length - 1].id,
        format: "romanLower",
        startValue: 1,
        includeInTotal: false,
      });
    }
    sections.push({
      ...createNumberingSection(firstNonToc.id, "Main Report"),
      endPageId: appendixStart.id === firstNonToc.id ? undefined : pages[Math.max(0, pages.findIndex((page) => page.id === appendixStart.id) - 1)]?.id,
      format: "decimal",
      startValue: 1,
    });
    if (appendixStart && appendixStart.id !== firstNonToc.id) {
      sections.push({
        ...createNumberingSection(appendixStart.id, "Appendix"),
        format: "decimal",
        prefix: "A-",
        startValue: 1,
      });
    }
    commitNavigation({ numberingSections: sections });
  }

  function updateManualTocEntry(entryId: string, patch: Partial<TocManualEntry>) {
    commitNavigation({
      manualTocEntries: manualTocEntries.map((entry) => (entry.id === entryId ? { ...entry, ...patch, updatedAt: new Date().toISOString() } : entry)),
    });
  }

  function getPageInsertionIndex(position: InsertPosition) {
    if (position === "start") return 0;
    if (position === "end") return pages.length;
    const activeIds = getActivePageIds();
    const indexes = pages.map((page, index) => (activeIds.includes(page.id) ? index : -1)).filter((index) => index >= 0);
    if (indexes.length === 0) return position === "before" ? currentPage - 1 : currentPage;
    return position === "before" ? Math.min(...indexes) : Math.max(...indexes) + 1;
  }

  function getTocInsertionIndex(candidatePages = pages) {
    if (tocSettings.insertPosition === "beforeFirst") return 0;
    const activeIds = getActivePageIds();
    const indexes = candidatePages.map((page, index) => (activeIds.includes(page.id) ? index : -1)).filter((index) => index >= 0);
    return indexes.length > 0 ? Math.max(...indexes) + 1 : Math.min(currentPage, candidatePages.length);
  }

  async function handleImportPdf(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []);
    event.target.value = "";
    if (files.length === 0) return;
    await addFilesToMergeQueue(files);
    setActivePageDialog("merge");
  }

  async function addFilesToMergeQueue(files: File[]) {
    const items: MergeQueueItem[] = [];
    for (const file of files) {
      const id = crypto.randomUUID();
      try {
        if (file.size === 0) throw new Error("Empty file.");
        if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) throw new Error("Unsupported file type.");
        const bytes = new Uint8Array(await file.arrayBuffer());
        const pdfDoc = await getDocument({ data: bytes.slice() }).promise;
        items.push({
          id,
          file,
          name: file.name,
          bytes,
          sourceDocumentId: crypto.randomUUID(),
          pdfDoc,
          pageCount: pdfDoc.numPages,
          range: `1-${pdfDoc.numPages}`,
        });
      } catch (error) {
        items.push({
          id,
          file,
          name: file.name,
          bytes: new Uint8Array(),
          sourceDocumentId: crypto.randomUUID(),
          pdfDoc: null as unknown as PDFDocumentProxy,
          pageCount: 0,
          range: "",
          error: getPdfErrorMessage(error, "upload"),
        });
      }
    }
    setMergeQueue((existing) => [...existing, ...items]);
  }

  async function importPdfFiles(files: File[], position: "before" | "after" | "end" = "after") {
    try {
      setErrorMessage("");
      const importedPages: PageView[] = [];
      const importedSources: Record<string, SourceDocument> = {};
      for (const [fileIndex, file] of files.entries()) {
        if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
          setErrorMessage(`${file.name} is not a supported PDF file.`);
          continue;
        }
        setWorkState({ message: `Importing ${file.name}`, progress: (fileIndex / Math.max(1, files.length)) * 80 });
        const bytes = new Uint8Array(await file.arrayBuffer());
        const sourceId = crypto.randomUUID();
        importedSources[sourceId] = { id: sourceId, name: file.name, bytes, mimeType: "application/pdf" };
        const importedDoc = await getDocument({ data: bytes.slice() }).promise;
        for (let pageIndex = 1; pageIndex <= importedDoc.numPages; pageIndex += 1) {
          importedPages.push(await renderPage(await importedDoc.getPage(pageIndex), sourceId));
        }
      }
      if (importedPages.length === 0) {
        setWorkState(null);
        return;
      }
      setSourceDocuments((existing) => ({ ...existing, ...importedSources }));
      const insertIndex = getPageInsertionIndex(position);
      const nextPages = renumberPages([...pages.slice(0, insertIndex), ...importedPages, ...pages.slice(insertIndex)]);
      const selectedIds = importedPages.map((page) => page.id);
      commitDocument(nextPages, marks, null, insertIndex + 1, selectedIds);
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(getPdfErrorMessage(error, "upload"));
    }
  }

  async function importMergeQueue() {
    const validItems = mergeQueue.filter((item) => !item.error && item.pageCount > 0);
    if (validItems.length === 0) {
      setErrorMessage("Add at least one valid PDF before importing.");
      return;
    }

    try {
      setErrorMessage("");
      const importedPages: PageView[] = [];
      const importedSources: Record<string, SourceDocument> = {};
      let completed = 0;
      const totalPages = validItems.reduce((total, item) => total + parsePageRanges(item.range, item.pageCount).flat().length, 0);

      for (const item of validItems) {
        const ranges = parsePageRanges(item.range || `1-${item.pageCount}`, item.pageCount);
        importedSources[item.sourceDocumentId] = { id: item.sourceDocumentId, name: item.name, bytes: item.bytes, mimeType: "application/pdf" };
        for (const pageNumber of ranges.flat()) {
          completed += 1;
          setWorkState({ message: `Importing ${item.name}`, progress: (completed / Math.max(1, totalPages)) * 90 });
          importedPages.push(await renderPage(await item.pdfDoc.getPage(pageNumber), item.sourceDocumentId));
        }
      }

      const insertIndex = getPageInsertionIndex(mergeInsertPosition);
      const nextPages = renumberPages([...pages.slice(0, insertIndex), ...importedPages, ...pages.slice(insertIndex)]);
      setSourceDocuments((existing) => ({ ...existing, ...importedSources }));
      commitDocument(nextPages, marks, null, insertIndex + 1, importedPages.map((page) => page.id));
      setMergeQueue([]);
      setActivePageDialog(null);
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(error instanceof Error ? error.message : "The selected PDFs could not be imported.");
    }
  }

  function moveMergeQueueItem(id: string, direction: -1 | 1) {
    setMergeQueue((existing) => {
      const index = existing.findIndex((item) => item.id === id);
      const nextIndex = index + direction;
      if (index < 0 || nextIndex < 0 || nextIndex >= existing.length) return existing;
      const next = [...existing];
      const [item] = next.splice(index, 1);
      next.splice(nextIndex, 0, item);
      return next;
    });
  }

  function updateMergeQueueItem(id: string, patch: Partial<MergeQueueItem>) {
    setMergeQueue((existing) => existing.map((item) => (item.id === id ? { ...item, ...patch } : item)));
  }

  function removeMergeQueueItem(id: string) {
    setMergeQueue((existing) => existing.filter((item) => item.id !== id));
  }

  async function handleReplaceFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      setErrorMessage("");
      const bytes = new Uint8Array(await file.arrayBuffer());
      const pdfDoc = await getDocument({ data: bytes.slice() }).promise;
      setReplaceFile({
        id: crypto.randomUUID(),
        file,
        name: file.name,
        bytes,
        sourceDocumentId: crypto.randomUUID(),
        pdfDoc,
        pageCount: pdfDoc.numPages,
        range: `1-${pdfDoc.numPages}`,
      });
      setReplaceSourcePage(1);
    } catch (error) {
      setErrorMessage(getPdfErrorMessage(error, "upload"));
    }
  }

  async function replaceSelectedPage() {
    const destination = pages.find((page) => page.id === getActivePageIds()[0]);
    if (!destination || !replaceFile || replaceFile.error) return;
    try {
      setErrorMessage("");
      setWorkState({ message: "Replacing page", progress: 30 });
      const sourcePage = await replaceFile.pdfDoc.getPage(clamp(Math.round(replaceSourcePage), 1, replaceFile.pageCount));
      const replacement = await renderPage(sourcePage, replaceFile.sourceDocumentId);
      const nextPage: PageView = {
        ...replacement,
        id: destination.id,
        pageNumber: destination.pageNumber,
        label: destination.label,
      };
      const nextPages = pages.map((page) => (page.id === destination.id ? nextPage : page));
      const nextMarks = replacePreserveAnnotations
        ? scaleMarksForReplacement(marks, destination, nextPage)
        : marks.filter((mark) => markPageId(mark) !== destination.id);
      setSourceDocuments((existing) => ({ ...existing, [replaceFile.sourceDocumentId]: { id: replaceFile.sourceDocumentId, name: replaceFile.name, bytes: replaceFile.bytes, mimeType: "application/pdf" } }));
      commitDocument(nextPages, nextMarks, null, destination.pageNumber, [destination.id]);
      setReplaceFile(null);
      setActivePageDialog(null);
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(getPdfErrorMessage(error, "render"));
    }
  }

  function hasPdfFiles(files: FileList) {
    return Array.from(files).some((file) => file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf"));
  }

  function hasDocxFiles(files: FileList) {
    return Array.from(files).some((file) => file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || file.type === "application/vnd.openxmlformats-officedocument.presentationml.presentation" || file.name.toLowerCase().endsWith(".docx") || file.name.toLowerCase().endsWith(".doc") || file.name.toLowerCase().endsWith(".pptx") || file.name.toLowerCase().endsWith(".ppt"));
  }

  function handlePagesPanelDragOver(event: ReactDragEvent<HTMLDivElement>) {
    if (!hasPdfFiles(event.dataTransfer.files) && !hasDocxFiles(event.dataTransfer.files)) return;
    event.preventDefault();
  }

  async function handlePagesPanelDrop(event: ReactDragEvent<HTMLDivElement>) {
    const docxFile = Array.from(event.dataTransfer.files).find((file) => file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || file.name.toLowerCase().endsWith(".docx") || file.name.toLowerCase().endsWith(".doc"));
    if (docxFile) {
      event.preventDefault();
      beginDocxImport(docxFile, hasOpenProject ? "append" : "replace");
      return;
    }
    const files = Array.from(event.dataTransfer.files).filter((file) => file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf"));
    if (files.length === 0) return;
    event.preventDefault();
    await importPdfFiles(files, "after");
  }

  function copySelectedPages() {
    const activeIds = getActivePageIds();
    const copied = pages.filter((page) => activeIds.includes(page.id));
    setCopiedPages({ pages: copied, marks: marksForPageIds(copied.map((page) => page.id)) });
  }

  function pasteCopiedPages(position: "before" | "after" | "end" = "after") {
    if (!copiedPages || copiedPages.pages.length === 0) return;
    const insertIndex = position === "end" ? pages.length : getPageInsertionIndex(position);
    const oldToNewPageId = new Map<string, string>();
    const pageCopies = copiedPages.pages.map((page, index) => {
      const nextId = crypto.randomUUID();
      oldToNewPageId.set(page.id, nextId);
      return { ...page, id: nextId, pageNumber: insertIndex + index + 1, label: page.label ? `${page.label} Copy` : page.label };
    });
    const markCopies = copiedPages.marks.map((mark): Mark => ({
      ...duplicateMark(mark, pageCopies.find((page) => page.id === oldToNewPageId.get(mark.pageId ?? ""))?.pageNumber ?? mark.page),
      pageId: mark.pageId ? oldToNewPageId.get(mark.pageId) : undefined,
      comment: mark.comment ? cloneComment(mark.comment) : mark.comment,
    })).filter((mark): mark is Mark => Boolean(mark.pageId));
    const nextPages = renumberPages([...pages.slice(0, insertIndex), ...pageCopies, ...pages.slice(insertIndex)]);
    const nextMarks = [...marks, ...markCopies];
    commitDocument(nextPages, nextMarks, null, insertIndex + 1, pageCopies.map((page) => page.id));
  }

  async function extractSelectedPages(removeAfterExtract: boolean) {
    const activeIds = getActivePageIds();
    const selectedPages = pages.filter((page) => activeIds.includes(page.id));
    if (selectedPages.length === 0) return;
    const exportedPdf = await buildPdfFromPages(selectedPages, marksForPageIds(selectedPages.map((page) => page.id)));
    downloadPdfBytes(await exportedPdf.save(), `${pdfName.replace(/\.pdf$/i, "")}-extract.pdf`);
    if (removeAfterExtract) deleteSelectedPages();
  }

  async function runExtractDialog() {
    const activeIds = getActivePageIds();
    const selectedPages = pages.filter((page) => activeIds.includes(page.id));
    if (selectedPages.length === 0) {
      setErrorMessage("Select at least one page to extract.");
      return;
    }
    try {
      const exportedPdf = await buildPdfFromPages(selectedPages, marksForPageIds(selectedPages.map((page) => page.id)));
      downloadPdfBytes(await exportedPdf.save(), `${(extractFilename.trim() || pdfName.replace(/\.pdf$/i, "") + "-extract").replace(/\.pdf$/i, "")}.pdf`);
      if (extractDeleteAfter) deleteSelectedPages();
      setActivePageDialog(null);
    } catch (error) {
      setErrorMessage(getPdfErrorMessage(error, "export"));
    }
  }

  async function splitAfterCurrentPage() {
    const first = pages.slice(0, currentPage);
    const second = pages.slice(currentPage);
    if (first.length === 0 || second.length === 0) {
      setErrorMessage("Choose a page that leaves pages on both sides of the split.");
      return;
    }
    downloadPdfBytes(await (await buildPdfFromPages(first, marksForPageIds(first.map((page) => page.id)))).save(), `${pdfName.replace(/\.pdf$/i, "")}-part-1.pdf`);
    downloadPdfBytes(await (await buildPdfFromPages(second, marksForPageIds(second.map((page) => page.id)))).save(), `${pdfName.replace(/\.pdf$/i, "")}-part-2.pdf`);
  }

  async function runSplitDialog() {
    try {
      const groups = getSplitGroups();
      if (groups.length === 0) {
        setErrorMessage("The split settings did not produce any output files.");
        return;
      }
      const base = pdfName.replace(/\.pdf$/i, "") || "publish-pro";
      for (const [index, group] of groups.entries()) {
        const outputPages = group.map((pageNumber) => pages[pageNumber - 1]).filter(Boolean);
        if (outputPages.length === 0) continue;
        const pdf = await buildPdfFromPages(outputPages, marksForPageIds(outputPages.map((page) => page.id)));
        downloadPdfBytes(await pdf.save(), `${base}-split-${index + 1}.pdf`);
      }
      setActivePageDialog(null);
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : "The document could not be split.");
    }
  }

  function getSplitGroups() {
    if (splitMode === "afterCurrent") {
      if (currentPage >= pages.length) throw new Error("Choose a current page before the final page.");
      return [pages.slice(0, currentPage).map((page) => page.pageNumber), pages.slice(currentPage).map((page) => page.pageNumber)];
    }
    if (splitMode === "everyN") {
      const size = clamp(Math.round(splitEveryN), 1, pages.length);
      const groups: number[][] = [];
      for (let index = 0; index < pages.length; index += size) groups.push(pages.slice(index, index + size).map((page) => page.pageNumber));
      return groups;
    }
    if (splitMode === "ranges") return parsePageRanges(splitRanges, pages.length);
    if (splitMode === "selected") {
      const selectedPages = pages.filter((page) => getActivePageIds().includes(page.id)).map((page) => page.pageNumber);
      return selectedPages.length ? [selectedPages] : [];
    }
    const boundaryIndexes = pages
      .map((page, index) => (getActivePageIds().includes(page.id) ? index : -1))
      .filter((index) => index >= 0)
      .sort((left, right) => left - right);
    if (boundaryIndexes.length === 0) return [];
    const groups: number[][] = [];
    let start = 0;
    for (const boundary of boundaryIndexes) {
      if (boundary >= start) {
        groups.push(pages.slice(start, boundary + 1).map((page) => page.pageNumber));
        start = boundary + 1;
      }
    }
    if (start < pages.length) groups.push(pages.slice(start).map((page) => page.pageNumber));
    return groups.filter((group) => group.length > 0);
  }

  function updateSelectedPageLabel(label: string) {
    const activeIds = getActivePageIds();
    const nextPages = pages.map((page) => (activeIds.includes(page.id) ? { ...page, label: label.trim() || undefined } : page));
    commitDocument(nextPages, marks, selectedMark, currentPage, activeIds);
  }

  function applyLabelPattern() {
    const activeIds = getActivePageIds();
    const selectedPages = pages.filter((page) => activeIds.includes(page.id));
    if (selectedPages.length === 0) return;
    const nextPages = pages.map((page) => {
      const index = selectedPages.findIndex((item) => item.id === page.id);
      if (index < 0) return page;
      return {
        ...page,
        label: `${labelPrefix}${formatSequenceValue(labelStart + index, labelFormat, labelPadding)}${labelSuffix}` || undefined,
      };
    });
    commitDocument(nextPages, marks, selectedMark, currentPage, activeIds);
    setActivePageDialog(null);
  }

  function startPageDrag(page: PageView) {
    const activeIds = selectedPageIds.includes(page.id) ? selectedPageIds : [page.id];
    setDraggedPageIds(activeIds);
  }

  function updatePageDropTarget(page: PageView, event: ReactDragEvent<HTMLButtonElement>) {
    event.preventDefault();
    const bounds = event.currentTarget.getBoundingClientRect();
    setPageDropTargetId(page.id);
    setPageDropPosition(event.clientY < bounds.top + bounds.height / 2 ? "before" : "after");
  }

  function finishPageDrop(targetPage: PageView) {
    if (draggedPageIds.length === 0 || draggedPageIds.includes(targetPage.id)) {
      setDraggedPageIds([]);
      setPageDropTargetId(null);
      return;
    }
    const moving = pages.filter((page) => draggedPageIds.includes(page.id));
    const remaining = pages.filter((page) => !draggedPageIds.includes(page.id));
    const targetIndexInRemaining = remaining.findIndex((page) => page.id === targetPage.id);
    const insertIndex = pageDropPosition === "before" ? targetIndexInRemaining : targetIndexInRemaining + 1;
    const nextPages = renumberPages([...remaining.slice(0, insertIndex), ...moving, ...remaining.slice(insertIndex)]);
    const nextMarks = marks;
    const firstMoved = moving[0];
    const nextCurrentPage = firstMoved ? nextPages.find((page) => page.id === firstMoved.id)?.pageNumber ?? currentPage : currentPage;
    commitDocument(nextPages, nextMarks, selectedMark, nextCurrentPage, moving.map((page) => page.id));
    setDraggedPageIds([]);
    setPageDropTargetId(null);
  }

  function commitMarks(nextMarks: Mark[], nextSelectedMark: string | null, historyBefore = marks) {
    commitDocument(pages, nextMarks, nextSelectedMark, currentPage, selectedPageIds, pages, historyBefore);
  }

  function commitDocument(
    nextPages: PageView[],
    nextMarks: Mark[],
    nextSelectedMark: string | null,
    nextCurrentPage: number,
    nextSelectedPageIds: string[],
    pagesBefore = pages,
    marksBefore = marks
  ) {
    const normalizedNextMarks = syncMarksToPages(migrateMarksToPageIds(nextMarks, pagesBefore), nextPages);
    if (
      arePagesEqual(pagesBefore, nextPages) &&
      areMarksEqual(marksBefore, normalizedNextMarks) &&
      selectedMark === nextSelectedMark &&
      currentPage === nextCurrentPage &&
      arraysEqual(selectedPageIds, nextSelectedPageIds)
    ) {
      return;
    }

    setHistory((existing) => ({
      past: [
        ...existing.past.slice(-(MAX_HISTORY_ENTRIES - 1)),
        {
          pagesBefore,
          pagesAfter: nextPages,
          marksBefore,
          marksAfter: normalizedNextMarks,
          selectedBefore: selectedMark,
          selectedAfter: nextSelectedMark,
          currentPageBefore: currentPage,
          currentPageAfter: nextCurrentPage,
          selectedPageIdsBefore: selectedPageIds,
          selectedPageIdsAfter: nextSelectedPageIds,
        },
      ],
      future: [],
    }));
    setPages(nextPages);
    setMarks(normalizedNextMarks);
    setSelectedMark(nextSelectedMark);
    setCurrentPage(nextCurrentPage);
    setSelectedPageIds(nextSelectedPageIds);
    setLastSelectedPageId(nextSelectedPageIds[0] ?? null);
  }

  function undo() {
    if (history.past.length === 0 && publishingHistory.past.length === 0 && navigationHistory.past.length > 0) {
      setNavigationHistory((existing) => {
        const previous = existing.past[existing.past.length - 1];
        if (!previous) return existing;
        const current = getNavigationSnapshot();
        setBookmarks(previous.bookmarks);
        setTocSettings(previous.tocSettings);
        setManualTocEntries(previous.manualTocEntries);
        setGeneratedToc(previous.generatedToc);
        setNumberingSections(previous.numberingSections ?? []);
        return {
          past: existing.past.slice(0, -1),
          future: [current, ...existing.future],
        };
      });
      return;
    }
    if (history.past.length === 0 && publishingHistory.past.length > 0) {
      setPublishingHistory((existing) => {
        const previous = existing.past[existing.past.length - 1];
        if (!previous) return existing;
        setPublishingSettings(previous);
        return {
          past: existing.past.slice(0, -1),
          future: [publishingSettings, ...existing.future],
        };
      });
      return;
    }
    setHistory((existing) => {
      const entry = existing.past[existing.past.length - 1];
      if (!entry) return existing;

      setPages(entry.pagesBefore);
      setMarks(entry.marksBefore);
      setSelectedMark(entry.selectedBefore);
      setCurrentPage(entry.currentPageBefore);
      setSelectedPageIds(entry.selectedPageIdsBefore);
      setLastSelectedPageId(entry.selectedPageIdsBefore[0] ?? null);

      return {
        past: existing.past.slice(0, -1),
        future: [entry, ...existing.future],
      };
    });
  }

  function redo() {
    if (history.future.length === 0 && publishingHistory.future.length === 0 && navigationHistory.future.length > 0) {
      setNavigationHistory((existing) => {
        const next = existing.future[0];
        if (!next) return existing;
        const current = getNavigationSnapshot();
        setBookmarks(next.bookmarks);
        setTocSettings(next.tocSettings);
        setManualTocEntries(next.manualTocEntries);
        setGeneratedToc(next.generatedToc);
        setNumberingSections(next.numberingSections ?? []);
        return {
          past: [...existing.past, current].slice(-MAX_HISTORY_ENTRIES),
          future: existing.future.slice(1),
        };
      });
      return;
    }
    if (history.future.length === 0 && publishingHistory.future.length > 0) {
      setPublishingHistory((existing) => {
        const next = existing.future[0];
        if (!next) return existing;
        setPublishingSettings(next);
        return {
          past: [...existing.past, publishingSettings].slice(-50),
          future: existing.future.slice(1),
        };
      });
      return;
    }
    setHistory((existing) => {
      const entry = existing.future[0];
      if (!entry) return existing;

      setPages(entry.pagesAfter);
      setMarks(entry.marksAfter);
      setSelectedMark(entry.selectedAfter);
      setCurrentPage(entry.currentPageAfter);
      setSelectedPageIds(entry.selectedPageIdsAfter);
      setLastSelectedPageId(entry.selectedPageIdsAfter[0] ?? null);

      return {
        past: [...existing.past, entry].slice(-MAX_HISTORY_ENTRIES),
        future: existing.future.slice(1),
      };
    });
  }

  async function buildPdfFromPages(
    exportPages: PageView[],
    exportMarks: Mark[],
    onProgress?: (message: string, progress: number) => void,
    publishingTargetContext: { currentPage?: number; selectedPageIds?: string[] } = {}
  ) {
    const exportedPdf = await PDFDocument.create();
    const loadedSourcePdfs = new Map<string, PDFDocument>();
    async function getSourcePdf(sourceDocumentId: string) {
      const existing = loadedSourcePdfs.get(sourceDocumentId);
      if (existing) return existing;
      const source = sourceDocuments[sourceDocumentId];
      if (!source) return null;
      if (source.mimeType && source.mimeType !== "application/pdf") return null;
      const loaded = await PDFDocument.load(source.bytes);
      loadedSourcePdfs.set(sourceDocumentId, loaded);
      return loaded;
    }

    const helvetica = await exportedPdf.embedFont(StandardFonts.Helvetica);
    const helveticaBold = await exportedPdf.embedFont(StandardFonts.HelveticaBold);
    const textFonts = new Map<string, PDFFont>();
    async function getTextBoxFont(style: TextBoxStyle) {
      const fontName = getStandardFontName(style);
      const existing = textFonts.get(fontName);
      if (existing) return existing;
      const embedded = await exportedPdf.embedFont(fontName);
      textFonts.set(fontName, embedded);
      return embedded;
    }

    for (const [pageIndex, pageView] of exportPages.entries()) {
      onProgress?.("Copying pages", 10 + (pageIndex / Math.max(1, exportPages.length)) * 20);
      if (pageView.sourceDocumentId && pageView.sourcePageNumber) {
        const sourcePdf = await getSourcePdf(pageView.sourceDocumentId);
        if (!sourcePdf) {
          exportedPdf.addPage([pageView.width / PAGE_SCALE, pageView.height / PAGE_SCALE]);
          continue;
        }
        const [copiedPage] = await exportedPdf.copyPages(sourcePdf, [pageView.sourcePageNumber - 1]);
        if (pageView.rotation) copiedPage.setRotation(degrees(pageView.rotation));
        exportedPdf.addPage(copiedPage);
      } else {
        const page = exportedPdf.addPage([pageView.width / PAGE_SCALE, pageView.height / PAGE_SCALE]);
        if (pageView.background) {
          page.drawRectangle({
            x: 0,
            y: 0,
            width: page.getWidth(),
            height: page.getHeight(),
            color: colorToRgb(pageView.background),
          });
        }
        if (pageView.rotation) page.setRotation(degrees(pageView.rotation));
      }
    }

    await drawPublishingMarksToPdf({
      pdf: exportedPdf,
      pages: exportPages,
      settings: publishingSettings,
      tokenContext: getPublishingTokenContext(),
      currentPage: publishingTargetContext.currentPage ?? currentPage,
      selectedPageIds: publishingTargetContext.selectedPageIds ?? selectedPageIds,
      imageAssets: publishingImageAssets,
      resolvedPageNumbers,
      layer: "watermark",
    });

    const pageIdToExportIndex = new Map(exportPages.map((page, index) => [page.id, index]));
    const tocExport = await drawGeneratedTocToPdf(exportedPdf, exportPages, pageIdToExportIndex);
    if (tocExport.lines.length > 0) {
      addTocLinks(exportedPdf, tocExport.tocPages, exportedPdf.getPages(), tocExport.lines, pageIdToExportIndex);
    }
    for (const [index, mark] of exportMarks.entries()) {
      onProgress?.("Applying annotations", 30 + (index / Math.max(1, exportMarks.length)) * 55);
      const exportPageId = markPageId(mark, exportPages);
      const exportIndex = exportPageId ? pageIdToExportIndex.get(exportPageId) : undefined;
      if (exportIndex === undefined) continue;
      const page = exportedPdf.getPage(exportIndex);
      const sourcePage = exportPages[exportIndex];
      if (!sourcePage) continue;

      const { width, height } = page.getSize();
      const x = (mark.x / sourcePage.width) * width;
      const y = height - ((mark.y + mark.height) / sourcePage.height) * height;
      const markWidth = (mark.width / sourcePage.width) * width;
      const markHeight = (mark.height / sourcePage.height) * height;
      const markSize = Math.max(8, (mark.size / sourcePage.height) * height * 1.2);

      if (mark.kind === "text") {
        const textStyle = getTextBoxStyle(mark);
        const textFont = await getTextBoxFont(textStyle);
        drawTextBoxToPdfPage({
          page,
          mark,
          font: textFont,
          x,
          y,
          width: markWidth,
          height: markHeight,
          fontSize: markSize,
          scale: height / sourcePage.height,
        });
        continue;
      }

      if (mark.kind === "shape" && mark.shapeStyle) {
        drawShapeToPdfPage({
          page,
          mark,
          font: helvetica,
          x,
          y,
          width: markWidth,
          height: markHeight,
          scale: height / sourcePage.height,
        });
        continue;
      }

      if (mark.kind === "comment" && mark.comment) {
        const markerSize = Math.max(14, (24 / sourcePage.height) * height);
        page.drawRectangle({
          x,
          y: y + markHeight - markerSize,
          width: markerSize,
          height: markerSize,
          color: colorToRgb(mark.comment.color),
          opacity: mark.comment.resolved ? 0.35 : 0.92,
          borderColor: colorToRgb("#7c2d12"),
          borderWidth: 0.75,
        });
        page.drawText("C", {
          x: x + markerSize * 0.28,
          y: y + markHeight - markerSize * 0.75,
          size: markerSize * 0.58,
          font: helveticaBold,
          color: colorToRgb("#111827"),
          opacity: mark.comment.resolved ? 0.65 : 1,
        });
        continue;
      }

      if ((mark.kind === "image" || mark.kind === "pngSignature") && mark.imageDataUrl) {
        const imageBytes = dataUrlToBytes(mark.imageDataUrl);
        const embeddedImage =
          mark.imageMimeType === "image/jpeg" ? await exportedPdf.embedJpg(imageBytes) : await exportedPdf.embedPng(imageBytes);
        const rotation = degreesToRadians(mark.rotation);
        const centerX = x + markWidth / 2;
        const centerY = y + markHeight / 2;
        const rotatedX = centerX - (markWidth / 2) * Math.cos(rotation) + (markHeight / 2) * Math.sin(rotation);
        const rotatedY = centerY - (markWidth / 2) * Math.sin(rotation) - (markHeight / 2) * Math.cos(rotation);

        page.drawImage(embeddedImage, {
          x: rotatedX,
          y: rotatedY,
          width: markWidth,
          height: markHeight,
          rotate: degrees(mark.rotation),
          opacity: mark.opacity ?? 1,
        });
        continue;
      }

      if (mark.kind === "stroke" && mark.strokePoints && mark.strokePoints.length > 0) {
        const exportedPoints = smoothStrokePoints(mark.strokePoints).map((point) => ({
          x: (point.x / sourcePage.width) * width,
          y: height - (point.y / sourcePage.height) * height,
        }));
        const strokeWidth = Math.max(0.5, (mark.size / sourcePage.height) * height);

        if (exportedPoints.length === 1) {
          const point = exportedPoints[0];
          page.drawCircle({
            x: point.x,
            y: point.y,
            size: strokeWidth / 2,
            color: colorToRgb(mark.color),
            opacity: mark.strokeOpacity ?? 1,
          });
        }

        for (let pointIndex = 1; pointIndex < exportedPoints.length; pointIndex += 1) {
          page.drawLine({
            start: exportedPoints[pointIndex - 1],
            end: exportedPoints[pointIndex],
            thickness: strokeWidth,
            color: colorToRgb(mark.color),
            opacity: mark.strokeOpacity ?? 1,
          });
        }
        continue;
      }

      if (isTextMarkupKind(mark.kind) && mark.markupRects && mark.markupRects.length > 0) {
        for (const rect of mark.markupRects) {
          const exportedRect = scaleMarkupRect(rect, sourcePage, { width, height });
          if (mark.kind === "textHighlight") {
            page.drawRectangle({
              x: exportedRect.x,
              y: exportedRect.y,
              width: exportedRect.width,
              height: exportedRect.height,
              color: colorToRgb(mark.color),
              opacity: mark.opacity ?? 0.38,
            });
            continue;
          }

          const lineY =
            mark.kind === "underline"
              ? exportedRect.y + Math.max(1, exportedRect.height * 0.12)
              : exportedRect.y + Math.max(1, exportedRect.height * 0.52);
          page.drawLine({
            start: { x: exportedRect.x, y: lineY },
            end: { x: exportedRect.x + exportedRect.width, y: lineY },
            thickness: Math.max(0.5, ((mark.thickness ?? 2) / sourcePage.height) * height),
            color: colorToRgb(mark.color),
            opacity: mark.opacity ?? 1,
          });
        }
        continue;
      }

      if (mark.kind === "highlight") {
        page.drawRectangle({
          x,
          y,
          width: markWidth,
          height: markHeight,
          color: colorToRgb(mark.color),
          opacity: 0.45,
        });
        continue;
      }

      if (mark.kind === "stamp") {
        page.drawRectangle({
          x,
          y,
          width: markWidth,
          height: markHeight,
          borderColor: colorToRgb(mark.color),
          borderWidth: 2,
        });
      }

      page.drawText(mark.text || " ", {
        x: x + (mark.kind === "stamp" ? 10 : 0),
        y: y + Math.max(4, markHeight / 3),
        size: markSize,
        font: mark.kind === "stamp" ? helveticaBold : helvetica,
        color: colorToRgb(mark.color),
      });
    }

    await drawPublishingMarksToPdf({
      pdf: exportedPdf,
      pages: exportPages,
      settings: publishingSettings,
      tokenContext: getPublishingTokenContext(),
      currentPage: publishingTargetContext.currentPage ?? currentPage,
      selectedPageIds: publishingTargetContext.selectedPageIds ?? selectedPageIds,
      imageAssets: publishingImageAssets,
      resolvedPageNumbers,
      layer: "foreground",
    });

    appendCommentsSummary(exportedPdf, helvetica, helveticaBold, exportMarks.filter((mark) => mark.kind === "comment" && mark.comment));
    addPdfOutline(exportedPdf, bookmarks, pageIdToExportIndex);
    return exportedPdf;
  }

  async function drawGeneratedTocToPdf(pdf: PDFDocument, exportPages: PageView[], pageIdToExportIndex: Map<string, number>) {
    const tocPageViews = exportPages.filter((page) => page.generatedToc);
    if (tocPageViews.length === 0) return { tocPages: [], lines: [] };
    const nonTocPages = exportPages.filter((page) => !page.generatedToc).map(pageToReference);
    const layout = layoutToc(buildTocEntries(bookmarks, manualTocEntries, nonTocPages, tocSettings), nonTocPages, tocSettings, getResolvedPageTextMap(nonTocPages));
    const font = await pdf.embedFont(getTocStandardFontName(tocSettings.fontFamily, false));
    const boldFont = await pdf.embedFont(getTocStandardFontName(tocSettings.fontFamily, true));
    const tocPages = tocPageViews
      .map((page) => pageIdToExportIndex.get(page.id))
      .filter((index): index is number => index !== undefined)
      .map((index) => pdf.getPage(index));
    tocPages.forEach((page, index) => {
      page.drawText(index === 0 ? tocSettings.title : `${tocSettings.title} ${index + 1}`, {
        x: tocSettings.marginLeft,
        y: page.getHeight() - tocSettings.marginTop + tocSettings.fontSize * 0.6,
        size: tocSettings.fontSize * 1.7,
        font: boldFont,
        color: colorToRgb(tocSettings.color),
      });
    });
    for (const line of layout.pages.flatMap((page) => page.lines)) {
      const page = tocPages[line.pageIndex];
      if (!page) continue;
      const isBold = tocSettings.boldLevels.includes(line.level);
      const fontToUse = isBold ? boldFont : font;
      const y = line.y;
      page.drawText(line.text, {
        x: line.x,
        y,
        size: tocSettings.fontSize,
        font: fontToUse,
        color: colorToRgb(tocSettings.color),
      });
      if (line.pageText) {
        const pageTextWidth = font.widthOfTextAtSize(line.pageText, tocSettings.fontSize);
        const pageTextX = page.getWidth() - tocSettings.marginRight - pageTextWidth;
        if (tocSettings.dotLeaders) {
          const leaderStart = line.x + fontToUse.widthOfTextAtSize(line.text, tocSettings.fontSize) + 8;
          const leaderEnd = pageTextX - 8;
          if (leaderEnd > leaderStart) {
            page.drawLine({
              start: { x: leaderStart, y: y + 2 },
              end: { x: leaderEnd, y: y + 2 },
              thickness: 0.35,
              color: colorToRgb("#9ca3af"),
              dashArray: [1, 3],
            });
          }
        }
        page.drawText(line.pageText, {
          x: pageTextX,
          y,
          size: tocSettings.fontSize,
          font,
          color: colorToRgb(tocSettings.color),
        });
      }
    }
    return { tocPages, lines: layout.pages.flatMap((page) => page.lines) };
  }

  async function exportPdf() {
    try {
      setErrorMessage("");
      setWorkState({ message: "Preparing export", progress: 10 });
      const exportedPdf = await buildPdfFromPages(pages, marks, (message, progress) => setWorkState({ message, progress }));
      setWorkState({ message: "Saving PDF", progress: 90 });
      const bytes = await exportedPdf.save();
      const filename = pdfName.replace(/\.pdf$/i, "") + "-edited.pdf";
      if (runtime.isDesktop) {
        const result = await runtime.saveFile({ kind: "pdf", bytes, suggestedName: filename });
        if (result?.saved && result.path) {
          setLastExportPath(result.path);
          setProjectStatusMessage(`Published PDF to ${filename}.`);
        }
      } else {
        downloadPdfBytes(bytes, filename);
      }
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(getPdfErrorMessage(error, "export"));
    }
  }


  return (
    <main className={`app-shell ${isLeftPanelCollapsed ? "left-collapsed" : ""} ${isRightPanelCollapsed ? "right-collapsed" : ""}`}>
      <AppChrome
        logoSrc={BRAND_ICON_SRC}
        documentName={pdfName}
        saveStatus={saveStatus}
        isBusy={isBusy}
        hasUnsavedChanges={hasUnsavedChanges}
        themeMode={themeMode}
        resolvedTheme={resolvedTheme}
        onThemeChange={changeThemeMode}
        onNewProject={startNewProject}
        onOpenProject={() => void chooseProjectFile()}
        onOpen={() => void choosePdfFile()}
        onImportDocx={() => void chooseDocxFile()}
        onImportPptx={() => void choosePptxFile()}
        canSaveProject={hasOpenProject}
        canUndo={canUndo}
        canRedo={canRedo}
        onSaveProject={saveProject}
        onSaveProjectAs={saveProjectAs}
        onCloseProject={closeProject}
        onUndo={undo}
        onRedo={redo}
        onShowShortcuts={() => setIsShortcutDialogOpen(true)}
        onExport={() => void exportPdf()}
      />
      <input className="hidden-file-input" ref={projectInput} type="file" accept=".pproj,application/zip" onChange={handleProjectUpload} aria-label="Choose Publish Pro project file" />
      <input className="hidden-file-input" ref={fileInput} type="file" accept="application/pdf" onChange={handleUpload} aria-label="Choose PDF file" />
      <input className="hidden-file-input" ref={docxInput} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={handleDocxUpload} aria-label="Choose DOCX Word document" />
      <input className="hidden-file-input" ref={pptxInput} type="file" accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation" onChange={handlePptxUpload} aria-label="Choose PPTX PowerPoint presentation" />
      <input className="hidden-file-input" ref={importPdfInput} type="file" accept="application/pdf" multiple onChange={handleImportPdf} aria-label="Choose PDFs to import" />
      <input className="hidden-file-input" ref={imageInput} type="file" accept="image/png,image/jpeg,image/svg+xml,image/webp" onChange={handleImageUpload} aria-label="Choose image file" />
      <input className="hidden-file-input" ref={projectImageInput} type="file" accept="image/png,image/jpeg,image/svg+xml,image/webp" onChange={handleProjectImageAssetUpload} aria-label="Choose project image asset" />
      <input className="hidden-file-input" ref={signatureInput} type="file" accept="image/png" onChange={handleSignatureUpload} aria-label="Choose PNG signature file" />

      {isProjectDialogOpen ? (
        <div className="modal-backdrop" role="presentation">
          <section className="modal-card" role="dialog" aria-modal="true" aria-labelledby="new-project-title">
            <div className="modal-header">
              <h2 id="new-project-title">New Project</h2>
              <button className="panel-icon-button" onClick={() => setIsProjectDialogOpen(false)} aria-label="Close new project dialog" title="Close dialog">x</button>
            </div>
            <div className="dialog-body">
              <label>
                Project name
                <input value={newProjectDraft.name} onChange={(event) => setNewProjectDraft((draft) => ({ ...draft, name: event.currentTarget.value }))} />
              </label>
              <label>
                Client
                <input value={newProjectDraft.client} onChange={(event) => setNewProjectDraft((draft) => ({ ...draft, client: event.currentTarget.value }))} />
              </label>
              <label>
                Author
                <input value={newProjectDraft.author} onChange={(event) => setNewProjectDraft((draft) => ({ ...draft, author: event.currentTarget.value }))} />
              </label>
              <label>
                Description
                <textarea value={newProjectDraft.description} onChange={(event) => setNewProjectDraft((draft) => ({ ...draft, description: event.currentTarget.value }))} rows={3} />
              </label>
              <p className="dialog-note">Metadata can be edited later in Assemble workspace settings.</p>
              <div className="dialog-actions">
                <button className="button ghost" onClick={() => setIsProjectDialogOpen(false)}>Cancel</button>
                <button className="button ghost" onClick={() => createNewProjectFromDraft("pdf")}>Create from PDF</button>
                <button className="button ghost" onClick={() => { setIsProjectDialogOpen(false); void chooseDocxFile(); }}>Import Word document</button>
                <button className="button ghost" onClick={() => { setIsProjectDialogOpen(false); void choosePptxFile(); }}>Import PowerPoint presentation</button>
                <button className="button primary" onClick={() => createNewProjectFromDraft("blank")}>Create Empty Project</button>
              </div>
            </div>
          </section>
        </div>
      ) : null}

      {pendingDocxFile ? (
        <DocxImportDialog
          file={pendingDocxFile}
          mode={docxImportMode}
          options={docxImportOptions}
          reimportTarget={docxReimportTarget}
          reimportImpact={docxReimportImpact}
          hasOpenProject={hasOpenProject}
          isBusy={isBusy}
          onModeChange={setDocxImportMode}
          onOptionsChange={setDocxImportOptions}
          onReplacementFile={(file) => {
            try {
              validateDocxImportFile(file);
              setPendingDocxFile(file);
              setErrorMessage("");
            } catch (error) {
              setErrorMessage(error instanceof Error ? error.message : "Choose a valid replacement DOCX file.");
            }
          }}
          onCancel={() => { setPendingDocxFile(null); setDocxReimportTarget(null); }}
          onImport={() => void confirmDocxImport()}
        />
      ) : null}

      {docxImportReport ? (
        <DocxImportReportDialog report={docxImportReport} onClose={() => setDocxImportReport(null)} />
      ) : null}

      {pendingPptxFile ? (
        <PptxImportDialog
          file={pendingPptxFile}
          mode={pptxImportMode}
          options={pptxImportOptions}
          reimportTarget={pptxReimportTarget}
          reimportImpact={pptxReimportImpact}
          hasOpenProject={hasOpenProject}
          isBusy={isBusy}
          onModeChange={setPptxImportMode}
          onOptionsChange={setPptxImportOptions}
          onReplacementFile={(file) => {
            try {
              validatePptxImportFile(file);
              setPendingPptxFile(file);
              setErrorMessage("");
            } catch (error) {
              setErrorMessage(error instanceof Error ? error.message : "Choose a valid replacement PPTX file.");
            }
          }}
          onCancel={() => { setPendingPptxFile(null); setPptxReimportTarget(null); }}
          onImport={() => void confirmPptxImport()}
        />
      ) : null}

      {pptxImportReport ? (
        <PptxImportReportDialog report={pptxImportReport} onClose={() => setPptxImportReport(null)} />
      ) : null}

      {pageContextMenu ? (
        <div className="page-context-menu" role="menu" style={{ left: pageContextMenu.x, top: pageContextMenu.y }} aria-label="Page actions">
          <button role="menuitem" onClick={() => runPageMenuAction(() => { setActivePageDialog("blank"); })}>Insert Blank Page Before</button>
          <button role="menuitem" onClick={() => runPageMenuAction(() => { setActivePageDialog("blank"); })}>Insert Blank Page After</button>
          <button role="menuitem" onClick={() => runPageMenuAction(() => { setMergeInsertPosition("before"); setActivePageDialog("merge"); })}>Import PDF Before</button>
          <button role="menuitem" onClick={() => runPageMenuAction(() => { setMergeInsertPosition("after"); setActivePageDialog("merge"); })}>Import PDF After</button>
          <button role="menuitem" onClick={() => runPageMenuAction(duplicateSelectedPages)}>Duplicate</button>
          <button role="menuitem" onClick={() => runPageMenuAction(copySelectedPages)}>Copy</button>
          <button role="menuitem" disabled={!copiedPages} onClick={() => runPageMenuAction(() => pasteCopiedPages("before"))}>Paste Before</button>
          <button role="menuitem" disabled={!copiedPages} onClick={() => runPageMenuAction(() => pasteCopiedPages("after"))}>Paste After</button>
          <button role="menuitem" onClick={() => runPageMenuAction(() => rotateSelectedPages(-90))}>Rotate Left</button>
          <button role="menuitem" onClick={() => runPageMenuAction(() => rotateSelectedPages(90))}>Rotate Right</button>
          <button role="menuitem" onClick={() => runPageMenuAction(() => { setActivePageDialog("replace"); })}>Replace Page</button>
          <button role="menuitem" onClick={() => runPageMenuAction(() => { setActivePageDialog("extract"); })}>Extract Selected Pages</button>
          <button role="menuitem" disabled={currentPage >= pages.length} onClick={() => runPageMenuAction(() => splitAfterCurrentPage())}>Split After This Page</button>
          <button role="menuitem" onClick={() => runPageMenuAction(() => { setActivePageDialog("labels"); })}>Edit Page Label</button>
          <button role="menuitem" className="danger" onClick={() => runPageMenuAction(deleteSelectedPages)}>Delete</button>
        </div>
      ) : null}

      {activePageDialog ? (
        <div className="modal-backdrop" role="presentation">
          <section className="modal-card" role="dialog" aria-modal="true" aria-label="Page management dialog">
            <div className="modal-header">
              <h2>{getPageDialogTitle(activePageDialog)}</h2>
              <button className="panel-icon-button" onClick={() => setActivePageDialog(null)} aria-label="Close dialog" title="Close dialog">x</button>
            </div>

            {activePageDialog === "merge" ? (
              <div className="dialog-body">
                <label className="file-picker">
                  Add PDF files
                  <input type="file" accept="application/pdf" multiple onChange={(event) => void addFilesToMergeQueue(Array.from(event.currentTarget.files ?? []))} />
                </label>
                <label>
                  Insert position
                  <select value={mergeInsertPosition} onChange={(event) => setMergeInsertPosition(event.currentTarget.value as InsertPosition)}>
                    <option value="start">Beginning</option>
                    <option value="before">Before current page</option>
                    <option value="after">After current page</option>
                    <option value="end">End</option>
                  </select>
                </label>
                <div className="merge-queue">
                  {mergeQueue.map((item, index) => (
                    <div className={`merge-item ${item.error ? "invalid" : ""}`} key={item.id}>
                      <strong>{item.name}</strong>
                      <span>{item.error ?? `${item.pageCount} pages`}</span>
                      <label>
                        Pages
                        <input value={item.range} onChange={(event) => updateMergeQueueItem(item.id, { range: event.currentTarget.value })} disabled={Boolean(item.error)} />
                      </label>
                      <div className="dialog-actions compact">
                        <button className="button ghost" onClick={() => moveMergeQueueItem(item.id, -1)} disabled={index === 0}>Up</button>
                        <button className="button ghost" onClick={() => moveMergeQueueItem(item.id, 1)} disabled={index === mergeQueue.length - 1}>Down</button>
                        <button className="button ghost" onClick={() => removeMergeQueueItem(item.id)}>Remove</button>
                      </div>
                    </div>
                  ))}
                  {mergeQueue.length === 0 ? <p className="dialog-note">No files selected.</p> : null}
                </div>
                <p className="dialog-note">{mergeQueue.filter((item) => !item.error).reduce((total, item) => total + safeRangeCount(item.range, item.pageCount), 0)} pages ready to insert.</p>
                <div className="dialog-actions">
                  <button className="button ghost" onClick={() => setActivePageDialog(null)}>Cancel</button>
                  <button className="button primary" onClick={() => void importMergeQueue()} disabled={!mergeQueue.some((item) => !item.error)}>Import PDFs</button>
                </div>
              </div>
            ) : null}

            {activePageDialog === "blank" ? (
              <div className="dialog-body">
                <label>
                  Preset
                  <select value={blankPagePreset} onChange={(event) => setBlankPagePreset(event.currentTarget.value as BlankPagePreset)}>
                    <option value="matchCurrent">Match current page</option>
                    <option value="a4Portrait">A4</option>
                    <option value="a3Portrait">A3</option>
                    <option value="letterPortrait">Letter</option>
                    <option value="legalPortrait">Legal</option>
                    <option value="custom">Custom</option>
                  </select>
                </label>
                <label>
                  Orientation
                  <select value={blankPageOrientation} onChange={(event) => setBlankPageOrientation(event.currentTarget.value as PageOrientation)}>
                    <option value="portrait">Portrait</option>
                    <option value="landscape">Landscape</option>
                  </select>
                </label>
                <div className="page-field-grid">
                  <label>Width<input type="number" value={customPageWidth} onChange={(event) => setCustomPageWidth(Number(event.currentTarget.value))} disabled={blankPagePreset !== "custom"} /></label>
                  <label>Height<input type="number" value={customPageHeight} onChange={(event) => setCustomPageHeight(Number(event.currentTarget.value))} disabled={blankPagePreset !== "custom"} /></label>
                </div>
                <label>
                  Units
                  <select value={blankPageUnit} onChange={(event) => setBlankPageUnit(event.currentTarget.value as DimensionUnit)} disabled={blankPagePreset !== "custom"}>
                    <option value="mm">Millimetres</option>
                    <option value="cm">Centimetres</option>
                    <option value="in">Inches</option>
                    <option value="pt">Points</option>
                  </select>
                </label>
                <div className="page-field-grid">
                  <label>Quantity<input type="number" min="1" max="50" value={blankPageQuantity} onChange={(event) => setBlankPageQuantity(Number(event.currentTarget.value))} /></label>
                  <label>Label<input value={blankPageLabel} onChange={(event) => setBlankPageLabel(event.currentTarget.value)} /></label>
                </div>
                <label>Background<input type="color" value={blankPageBackground} onChange={(event) => setBlankPageBackground(event.currentTarget.value)} /></label>
                <div className="dialog-actions">
                  <button className="button ghost" onClick={() => setActivePageDialog(null)}>Cancel</button>
                  <button className="button ghost" onClick={() => { insertBlankPages("before"); setActivePageDialog(null); }}>Insert Before</button>
                  <button className="button primary" onClick={() => { insertBlankPages("after"); setActivePageDialog(null); }}>Insert After</button>
                </div>
              </div>
            ) : null}

            {activePageDialog === "replace" ? (
              <div className="dialog-body">
                <label className="file-picker">Replacement PDF<input type="file" accept="application/pdf" onChange={(event) => void handleReplaceFile(event)} /></label>
                {replaceFile ? <p className="dialog-note">{replaceFile.name} - {replaceFile.pageCount} pages</p> : null}
                <label>Source page<input type="number" min="1" max={replaceFile?.pageCount ?? 1} value={replaceSourcePage} onChange={(event) => setReplaceSourcePage(Number(event.currentTarget.value))} /></label>
                <label className="checkbox-row"><input type="checkbox" checked={replacePreserveAnnotations} onChange={(event) => setReplacePreserveAnnotations(event.currentTarget.checked)} />Preserve existing annotations proportionally</label>
                {replacePreserveAnnotations ? <p className="dialog-note">If page dimensions differ, annotation coordinates are scaled and clamped to the replacement page.</p> : null}
                <div className="dialog-actions">
                  <button className="button ghost" onClick={() => setActivePageDialog(null)}>Cancel</button>
                  <button className="button primary" onClick={() => void replaceSelectedPage()} disabled={!replaceFile}>Replace Page</button>
                </div>
              </div>
            ) : null}

            {activePageDialog === "split" ? (
              <div className="dialog-body">
                <label>Split method<select value={splitMode} onChange={(event) => setSplitMode(event.currentTarget.value as SplitMode)}><option value="afterCurrent">After current page</option><option value="everyN">Every N pages</option><option value="ranges">By page ranges</option><option value="selected">Selected pages only</option><option value="selectedBoundaries">At selected page boundaries</option></select></label>
                {splitMode === "everyN" ? <label>Pages per file<input type="number" min="1" max={pages.length} value={splitEveryN} onChange={(event) => setSplitEveryN(Number(event.currentTarget.value))} /></label> : null}
                {splitMode === "ranges" ? <label>Ranges<input value={splitRanges} onChange={(event) => setSplitRanges(event.currentTarget.value)} placeholder="1-5, 8-10" /></label> : null}
                <p className="dialog-note">Split output uses separate PDF downloads. ZIP packaging is not included yet.</p>
                <div className="dialog-actions"><button className="button ghost" onClick={() => setActivePageDialog(null)}>Cancel</button><button className="button primary" onClick={() => void runSplitDialog()}>Split PDF</button></div>
              </div>
            ) : null}

            {activePageDialog === "extract" ? (
              <div className="dialog-body">
                <label>Output filename<input value={extractFilename} onChange={(event) => setExtractFilename(event.currentTarget.value)} placeholder={`${pdfName.replace(/\.pdf$/i, "")}-extract.pdf`} /></label>
                <label className="checkbox-row"><input type="checkbox" checked={extractDeleteAfter} onChange={(event) => setExtractDeleteAfter(event.currentTarget.checked)} />Delete pages after successful extraction</label>
                <p className="dialog-note">{activePageIds.length} selected pages will be exported with their annotations and comment summary.</p>
                <div className="dialog-actions"><button className="button ghost" onClick={() => setActivePageDialog(null)}>Cancel</button><button className="button primary" onClick={() => void runExtractDialog()}>Extract</button></div>
              </div>
            ) : null}

            {activePageDialog === "labels" ? (
              <div className="dialog-body">
                <label>Prefix<input value={labelPrefix} onChange={(event) => setLabelPrefix(event.currentTarget.value)} placeholder="A-" /></label>
                <label>Suffix<input value={labelSuffix} onChange={(event) => setLabelSuffix(event.currentTarget.value)} placeholder="" /></label>
                <div className="page-field-grid"><label>Start<input type="number" value={labelStart} onChange={(event) => setLabelStart(Number(event.currentTarget.value))} /></label><label>Padding<input type="number" min="0" max="6" value={labelPadding} onChange={(event) => setLabelPadding(Number(event.currentTarget.value))} /></label></div>
                <label>Format<select value={labelFormat} onChange={(event) => setLabelFormat(event.currentTarget.value as "number" | "roman" | "alpha")}><option value="number">Number</option><option value="roman">Roman</option><option value="alpha">Alphabetic</option></select></label>
                <p className="dialog-note">Page labels are internal Publish Pro metadata and are shown in the workspace and comment exports.</p>
                <div className="dialog-actions"><button className="button ghost" onClick={() => setActivePageDialog(null)}>Cancel</button><button className="button primary" onClick={applyLabelPattern}>Apply Labels</button></div>
              </div>
            ) : null}
          </section>
        </div>
      ) : null}

      {isShortcutDialogOpen ? (
        <ShortcutReferenceDialog
          onClose={() => setIsShortcutDialogOpen(false)}
          isMac={navigator.platform.toLowerCase().includes("mac")}
        />
      ) : null}

      {isCloseGuardOpen ? (
        <div className="modal-backdrop" role="presentation">
          <section className="modal-card" role="dialog" aria-modal="true" aria-labelledby="close-guard-title">
            <div className="modal-header">
              <h2 id="close-guard-title">Unsaved changes</h2>
            </div>
            <div className="dialog-body">
              <p className="dialog-note">Save changes to {projectMetadata.name || "this project"} before closing Publish Pro?</p>
              <div className="dialog-actions">
                <button className="button ghost" onClick={() => setIsCloseGuardOpen(false)}>Cancel</button>
                <button className="button ghost" onClick={() => void discardAndCloseDesktop()}>Discard</button>
                <button className="button primary" onClick={() => void saveAndCloseDesktop()}>Save</button>
              </div>
            </div>
          </section>
        </div>
      ) : null}

      <nav className="main-toolbar" aria-label={`${activeWorkspaceItem.title} workspace tools`}>
        <div className="toolbar-context">
          <span>{activeWorkspaceItem.title}</span>
          <strong>{activeWorkspaceItem.description}</strong>
        </div>
        {activeWorkspace === "assemble" ? (
          <>
            <div className="toolbar-group" aria-label="Page actions">
              <ToolButton icon={<FilePlus2 />} label="Insert blank pages" onClick={() => setActivePageDialog("blank")} disabled={isBusy} />
              <ToolButton icon={<Files />} label="Combine files" onClick={() => setActivePageDialog("merge")} disabled={isBusy} />
            </div>
            <ToolbarMenu icon={<MoreHorizontal />} label="More page actions">
              <MenuToolButton icon={<PanelLeftOpen />} label="Select all pages" onClick={selectAllPages} disabled={isBusy} />
              <MenuToolButton icon={<PanelLeftClose />} label="Clear page selection" onClick={clearPageSelection} disabled={isBusy} />
              <MenuToolButton icon={<CornerUpRight />} label="Rotate left" onClick={() => void rotateSelectedPages(-90)} disabled={isBusy} />
              <MenuToolButton icon={<CornerUpRight />} label="Rotate 180" onClick={() => void rotateSelectedPages(180)} disabled={isBusy} />
              <MenuToolButton icon={<Files />} label="Replace page" onClick={() => setActivePageDialog("replace")} disabled={isBusy} />
              <MenuToolButton icon={<Download />} label="Extract pages" onClick={() => setActivePageDialog("extract")} disabled={isBusy} />
              <MenuToolButton icon={<Minus />} label="Split PDF" onClick={() => setActivePageDialog("split")} disabled={isBusy || pages.length <= 1} />
            </ToolbarMenu>
          </>
        ) : null}
        {activeWorkspace === "review" ? (
          <>
            <div className="toolbar-group" aria-label="Navigation and edit tools">
              <ToolButton icon={<MousePointer2 />} label="Select tool" active={activeTool === "select"} onClick={() => setActiveTool("select")} disabled={isBusy} />
              <ToolButton icon={<Hand />} label="Hand tool" disabled />
              <ToolButton icon={<ZoomOut />} label="Zoom out" onClick={() => setZoom((value) => Math.max(0.55, value - 0.1))} disabled={isBusy} />
              <span className="zoom-label">{Math.round(zoom * 100)}%</span>
              <ToolButton icon={<ZoomIn />} label="Zoom in" onClick={() => setZoom((value) => Math.min(1.6, value + 0.1))} disabled={isBusy} />
              <ToolButton icon={<Maximize2 />} label="Fit page" onClick={() => setZoom(1)} disabled={isBusy} />
            </div>
            <div className="toolbar-group" aria-label="Insert editable objects">
              <ToolButton icon={<TextCursorInput />} label="Add text" active={activeTool === "text"} onClick={() => setActiveTool("text")} disabled={isBusy} />
              <ToolButton icon={<ImageIcon />} label="Add image" active={activeTool === "image"} onClick={() => { setActiveTool("image"); imageInput.current?.click(); }} disabled={isBusy} />
              <ToolButton icon={<PenLine />} label="Upload PNG signature" active={activeTool === "pngSignature"} onClick={() => { setActiveTool("pngSignature"); signatureInput.current?.click(); }} disabled={isBusy} />
              <ToolButton icon={<PenLine />} label="Add signature text" active={activeTool === "signature"} onClick={() => setActiveTool("signature")} disabled={isBusy} />
            </div>
          </>
        ) : null}
        {activeWorkspace === "review" ? (
          <>
            <div className="toolbar-group" aria-label="Text markup tools">
              <ToolButton icon={<Highlighter />} label="Highlight selected text" active={activeTool === "textHighlight"} onClick={() => setActiveTool("textHighlight")} disabled={isBusy} />
              <ToolButton icon={<Underline />} label="Underline selected text" active={activeTool === "underline"} onClick={() => setActiveTool("underline")} disabled={isBusy} />
              <ToolButton icon={<Strikethrough />} label="Strikethrough selected text" active={activeTool === "strikethrough"} onClick={() => setActiveTool("strikethrough")} disabled={isBusy} />
              <ToolButton icon={<Pencil />} label="Draw freehand" active={activeTool === "draw"} onClick={() => setActiveTool("draw")} disabled={isBusy} />
              <ToolButton icon={<MessageSquare />} label="Add comment" active={activeTool === "comment"} onClick={() => setActiveTool("comment")} disabled={isBusy} />
            </div>
            <ToolbarMenu icon={<Square />} label="Shapes">
              <MenuToolButton icon={<Highlighter />} label="Area highlight" active={activeTool === "highlight"} onClick={() => setActiveTool("highlight")} disabled={isBusy} />
              <MenuToolButton icon={<Square />} label="Rectangle" active={activeTool === "shape:rectangle"} onClick={() => setActiveTool("shape:rectangle")} disabled={isBusy} />
              <MenuToolButton icon={<Circle />} label="Ellipse" active={activeTool === "shape:ellipse"} onClick={() => setActiveTool("shape:ellipse")} disabled={isBusy} />
              <MenuToolButton icon={<Minus />} label="Line" active={activeTool === "shape:line"} onClick={() => setActiveTool("shape:line")} disabled={isBusy} />
              <MenuToolButton icon={<CornerUpRight />} label="Arrow" active={activeTool === "shape:arrow"} onClick={() => setActiveTool("shape:arrow")} disabled={isBusy} />
              <MenuToolButton icon={<CornerUpRight />} label="Double arrow" active={activeTool === "shape:doubleArrow"} onClick={() => setActiveTool("shape:doubleArrow")} disabled={isBusy} />
              <MenuToolButton icon={<Square />} label="Rounded rectangle" active={activeTool === "shape:roundedRectangle"} onClick={() => setActiveTool("shape:roundedRectangle")} disabled={isBusy} />
              <MenuToolButton icon={<Square />} label="Polygon" active={activeTool === "shape:polygon"} onClick={() => setActiveTool("shape:polygon")} disabled={isBusy} />
              <MenuToolButton icon={<Circle />} label="Cloud" active={activeTool === "shape:cloud"} onClick={() => setActiveTool("shape:cloud")} disabled={isBusy} />
              <MenuToolButton icon={<MessageSquare />} label="Callout" active={activeTool === "shape:callout"} onClick={() => setActiveTool("shape:callout")} disabled={isBusy} />
            </ToolbarMenu>
          </>
        ) : null}
        {activeWorkspace === "review" ? (
          <div className="toolbar-group" aria-label="Review tools">
            <ToolButton icon={<MessageSquare />} label="Add comment" active={activeTool === "comment"} onClick={() => setActiveTool("comment")} disabled={isBusy} />
            <ToolButton icon={<CheckCircle2 />} label="Show open comments" onClick={() => setCommentFilter("open")} disabled={isBusy} />
            <ToolButton icon={<MessageSquare />} label="Show resolved comments" onClick={() => setCommentFilter("resolved")} disabled={isBusy} />
            <ToolButton icon={<Stamp />} label="Add approval stamp" active={activeTool === "stamp"} onClick={() => setActiveTool("stamp")} disabled={isBusy} />
          </div>
        ) : null}
        {activeWorkspace === "import" ? (
          <div className="toolbar-group" aria-label="Assembly tools">
            <ToolButton icon={<Files />} label="Open PDF" onClick={() => void choosePdfFile()} disabled={isBusy} />
            <ToolButton icon={<Plus />} label="Import PDFs" onClick={() => void choosePdfImports()} disabled={isBusy} />
            <ToolButton icon={<Files />} label="Combine files" onClick={() => setActivePageDialog("merge")} disabled={isBusy} />
            <ToolButton icon={<BookOpen />} label="Bookmarks" onClick={() => setLeftPanel("bookmarks")} disabled={isBusy} />
            <ToolButton icon={<FilePlus2 />} label="Insert blank pages" onClick={() => setActivePageDialog("blank")} disabled={isBusy} />
          </div>
        ) : null}
        {activeWorkspace === "publish" ? (
          <div className="toolbar-group" aria-label="Publish tools">
            <ToolButton icon={<Download />} label="Export edited PDF" onClick={() => void exportPdf()} disabled={isBusy} />
            <ToolButton icon={<Undo2 />} label="Undo" onClick={undo} disabled={!canUndo} />
            <ToolButton icon={<Redo2 />} label="Redo" onClick={redo} disabled={!canRedo} />
          </div>
        ) : null}
      </nav>

      <section className="workspace">
        <nav className="left-nav workspace-nav" aria-label="Workspaces">
          <span className="workspace-nav-label">Workspace</span>
          {workspaceItems.map((workspace) => (
            <WorkspaceButton
              key={workspace.id}
              icon={workspace.icon}
              title={workspace.title}
              description={workspace.description}
              active={activeWorkspace === workspace.id && !isLeftPanelCollapsed}
              onClick={() => selectWorkspace(workspace.id)}
            />
          ))}
          <button
            className="rail-button rail-collapse"
            onClick={() => setIsLeftPanelCollapsed((value) => !value)}
            title={isLeftPanelCollapsed ? "Show left panel" : "Collapse left panel"}
            aria-label={isLeftPanelCollapsed ? "Show left panel" : "Collapse left panel"}
          >
            {isLeftPanelCollapsed ? <PanelLeftOpen size={18} /> : <PanelLeftClose size={18} />}
          </button>
        </nav>

        {!isLeftPanelCollapsed ? (
          <aside className="left-panel" onDragOver={leftPanel === "pages" ? handlePagesPanelDragOver : undefined} onDrop={leftPanel === "pages" ? handlePagesPanelDrop : undefined}>
            <div className="panel-header">
              <div>
                <h2>{getWorkspacePanelTitle(activeWorkspace)}</h2>
                <span>{getWorkspacePanelSubtitle(activeWorkspace, pages.length, selectedPageCount, pdfName, comments.length)}</span>
              </div>
              <button className="panel-icon-button" onClick={() => setIsLeftPanelCollapsed(true)} title="Collapse left panel" aria-label="Collapse left panel">
                <PanelLeftClose size={16} />
              </button>
            </div>

            {activeWorkspace === "assemble" ? (
              <>
                <div className="searchbox">
                  <Search size={15} />
                  <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Find page" aria-label="Find page" />
                </div>
                <div
                  className="page-list"
                  onClick={(event) => {
                    if (event.target === event.currentTarget) clearPageSelection();
                  }}
                >
                  {filteredPages.map((page) => (
                    <button
                      className={`thumb ${currentPage === page.pageNumber ? "current" : ""} ${selectedPageIds.includes(page.id) ? "selected" : ""} ${
                        pageDropTargetId === page.id ? `drop-${pageDropPosition}` : ""
                      }`}
                      key={page.id}
                      draggable={!isBusy}
                      onClick={(event) => handlePageSelect(page, event)}
                      onContextMenu={(event) => openPageContextMenu(page, event)}
                      onKeyDown={(event) => handlePageThumbKeyDown(page, event)}
                      onDragStart={() => startPageDrag(page)}
                      onDragOver={(event) => updatePageDropTarget(page, event)}
                      onDrop={() => finishPageDrop(page)}
                      onDragEnd={() => {
                        setDraggedPageIds([]);
                        setPageDropTargetId(null);
                      }}
                      title={`Page ${page.pageNumber}${page.label ? `: ${page.label}` : ""} - ${Math.round(page.width / PAGE_SCALE)} x ${Math.round(page.height / PAGE_SCALE)} pt - ${page.width >= page.height ? "Landscape" : "Portrait"}`}
                      aria-label={`Page ${page.pageNumber}${page.label ? `, ${page.label}` : ""}`}
                      aria-current={currentPage === page.pageNumber ? "page" : undefined}
                      aria-pressed={selectedPageIds.includes(page.id)}
                    >
                      <span className="thumb-grip" aria-hidden="true">::</span>
                      <div className="thumb-page">
                        {page.imageUrl ? <img src={page.imageUrl} alt="" /> : <FilePlus2 size={28} />}
                      </div>
                      <span>Page {page.pageNumber}</span>
                      {page.label ? <em className="page-label">{page.label}</em> : null}
                      <small className="page-meta">
                        {page.width >= page.height ? "Landscape" : "Portrait"} · {pageAnnotations.get(page.id) ?? 0} marks
                      </small>
                      {page.rotation ? <small className="page-rotation">{page.rotation}deg</small> : null}
                    </button>
                  ))}
                </div>
                <div className="page-management">
                  <div className="page-panel-toolbar">
                    <span>{selectedPageCount > 1 ? `${selectedPageCount} selected` : "Page tools"}</span>
                    <details className="panel-menu">
                      <summary title="Page actions" aria-label="Page actions">
                        <MoreHorizontal size={18} />
                      </summary>
                      <div className="panel-menu-list" role="menu">
                        <button onClick={selectAllPages}>Select all</button>
                        <button onClick={clearPageSelection}>Clear selection</button>
                        <button onClick={() => setActivePageDialog("blank")}>Insert blank pages</button>
                        <button onClick={() => setActivePageDialog("merge")}>Combine files</button>
                        <button onClick={duplicatePage}>Duplicate</button>
                        <button onClick={copySelectedPages}>Copy</button>
                        <button onClick={() => pasteCopiedPages("after")} disabled={!copiedPages}>Paste</button>
                        <button onClick={() => void rotateSelectedPages(90)}>Rotate right</button>
                        <button onClick={() => void rotateSelectedPages(-90)}>Rotate left</button>
                        <button onClick={() => void rotateSelectedPages(180)}>Rotate 180</button>
                        <button onClick={() => setActivePageDialog("replace")}>Replace</button>
                        <button onClick={() => setActivePageDialog("extract")}>Extract</button>
                        <button onClick={() => setActivePageDialog("split")} disabled={pages.length <= 1}>Split</button>
                        <button onClick={() => setActivePageDialog("labels")}>Label pattern</button>
                        <button className="danger" onClick={deleteSelectedPages}>Delete</button>
                      </div>
                    </details>
                  </div>
                  <label className="page-label-inline">
                    <span>Label</span>
                    <input
                      value={pages.find((page) => selectedPageIds.includes(page.id))?.label ?? ""}
                      onChange={(event) => updateSelectedPageLabel(event.target.value)}
                      placeholder="Cover, A-01..."
                      disabled={isBusy}
                    />
                  </label>
                </div>
              </>
            ) : null}

            {activeWorkspace === "review" ? (
              <div className="workspace-panel-content">
                <ToolCard icon={<MousePointer2 />} title="Select and move" description="Choose, move, resize, duplicate, or delete existing objects." onClick={() => setActiveTool("select")} />
                <ToolCard icon={<TextCursorInput />} title="Text" description="Place editable overlay text boxes on the page." onClick={() => setActiveTool("text")} />
                <ToolCard icon={<ImageIcon />} title="Image" description="Add PNG, JPG, SVG, or WebP artwork." onClick={() => { setActiveTool("image"); imageInput.current?.click(); }} />
                <ToolCard icon={<PenLine />} title="Signature" description="Place transparent PNG signatures or signature text." onClick={() => { setActiveTool("pngSignature"); signatureInput.current?.click(); }} />
              </div>
            ) : null}

            {activeWorkspace === "review" ? (
              <div className="workspace-panel-content">
                <ToolCard icon={<Highlighter />} title="Text markup" description="Highlight, underline, or strike selected PDF text." onClick={() => setActiveTool("textHighlight")} />
                <ToolCard icon={<Pencil />} title="Draw" description="Use freehand strokes for notes, sketches, and signatures." onClick={() => setActiveTool("draw")} />
                <ToolCard icon={<Square />} title="Shapes" description="Add rectangles, ellipses, lines, arrows, clouds, and callouts." onClick={() => setActiveTool("shape:rectangle")} />
                <ToolCard icon={<MessageSquare />} title="Sticky note" description="Place comments directly on the current page." onClick={() => setActiveTool("comment")} />
              </div>
            ) : null}

            {activeWorkspace === "import" ? (
              <div className="workspace-panel-content">
                <div className="project-quick-actions">
                  <button className="button ghost" onClick={startNewProject} disabled={isBusy}>New</button>
                  <button className="button ghost" onClick={() => void chooseProjectFile()} disabled={isBusy}>Open Project</button>
                  <button className="button primary" onClick={() => void saveProject()} disabled={isBusy || !hasOpenProject}>Save</button>
                </div>
                <ToolCard icon={<Files />} title="Open PDF" description="Open a PDF as the working document." onClick={() => void choosePdfFile()} />
                <ToolCard icon={<Files />} title="Import Word document" description="Convert DOCX content into Publish Pro pages while preserving the source file." onClick={() => void chooseDocxFile()} />
                <ToolCard icon={<Files />} title="Import PowerPoint presentation" description="Convert PPTX slides into Publish Pro pages while preserving the source file." onClick={() => void choosePptxFile()} />
                <ToolCard icon={<Plus />} title="Import PDFs" description="Append or insert PDF pages into this publication." onClick={() => void choosePdfImports()} />
                <ToolCard icon={<ImageIcon />} title="Add image asset" description="Import image artwork for publishing marks without placing it on a page." onClick={() => startProjectImageAssetImport({ kind: "projectAssets" })} />
                <ToolCard icon={<PenLine />} title="Add signature asset" description="Add a local transparent PNG signature." onClick={() => { setActiveTool("pngSignature"); signatureInput.current?.click(); }} />
                <ImportSourceManager
                  sourceDocuments={sourceDocuments}
                  pages={pages}
                  docxImportGroups={docxImportGroups}
                  pptxImportGroups={pptxImportGroups}
                  isBusy={isBusy}
                  onReviewPage={(page) => {
                    setCurrentPage(page.pageNumber);
                    setSelectedPageIds([page.id]);
                    selectWorkspace("assemble");
                  }}
                  onReimportDocx={beginDocxReimport}
                  onReimportPptx={beginPptxReimport}
                />
                <details className="publishing-section" open>
                  <summary>Bookmarks</summary>
                  <div className="project-quick-actions">
                    <button className="button ghost" onClick={addBookmarkFromCurrentPage} disabled={isBusy}>Current page</button>
                    <button className="button ghost" onClick={addChildBookmark} disabled={isBusy || !selectedBookmarkId}>Add child</button>
                    <button className="button ghost" onClick={addBookmarksFromSelectedPages} disabled={isBusy}>Selected pages</button>
                  </div>
                  <button className="button ghost full-width" onClick={addBookmarksFromPageLabels} disabled={isBusy}>Create from page labels</button>
                  <div className="navigation-preview-list" role="tree" aria-label="Bookmark hierarchy">
                    {visibleBookmarks.length > 0 ? visibleBookmarks.slice(0, 8).map((bookmark) => {
                      const issue = navigationIssues.find((item) => item.bookmarkId === bookmark.id);
                      return (
                        <button key={bookmark.id} className={`navigation-preview-row ${selectedBookmarkId === bookmark.id ? "active" : ""}`} style={{ paddingLeft: 8 + (bookmark.level - 1) * 14 }} onClick={() => navigateToBookmark(bookmark)} role="treeitem" aria-level={bookmark.level}>
                          {issue ? <AlertTriangle size={14} aria-hidden="true" /> : <BookOpen size={14} aria-hidden="true" />}
                          <span>{bookmark.title}</span>
                          <small>{formatBookmarkDestination(bookmark, pages)}</small>
                        </button>
                      );
                    }) : <p className="muted-text">No bookmarks yet.</p>}
                  </div>
                </details>
                <details className="publishing-section">
                  <summary>Table of Contents</summary>
                  <label>Title<input value={tocSettings.title} onChange={(event) => commitNavigation({ tocSettings: { ...tocSettings, title: event.currentTarget.value } })} /></label>
                  <div className="page-field-grid">
                    <label>Source<select value={tocSettings.source} onChange={(event) => commitNavigation({ tocSettings: { ...tocSettings, source: event.currentTarget.value as TocSettings["source"] } })}><option value="bookmarks">Bookmarks</option><option value="pageLabels">Page labels</option><option value="manual">Manual entries</option></select></label>
                    <label>Max depth<input type="number" min="1" max="8" value={tocSettings.maxDepth} onChange={(event) => commitNavigation({ tocSettings: { ...tocSettings, maxDepth: clamp(Number(event.currentTarget.value) || 1, 1, 8) } })} /></label>
                  </div>
                  <div className="page-field-grid">
                    <label>Font<select value={tocSettings.fontFamily} onChange={(event) => commitNavigation({ tocSettings: { ...tocSettings, fontFamily: event.currentTarget.value as TocSettings["fontFamily"] } })}><option value="Helvetica">Helvetica</option><option value="Times Roman">Times Roman</option><option value="Courier">Courier</option></select></label>
                    <label>Size<input type="number" min="8" max="18" value={tocSettings.fontSize} onChange={(event) => commitNavigation({ tocSettings: { ...tocSettings, fontSize: clamp(Number(event.currentTarget.value) || 11, 8, 18) } })} /></label>
                  </div>
                  <label>Insert<select value={tocSettings.insertPosition} onChange={(event) => commitNavigation({ tocSettings: { ...tocSettings, insertPosition: event.currentTarget.value as TocSettings["insertPosition"] } })}><option value="beforeFirst">Before first page</option><option value="afterSelected">After selected page</option></select></label>
                  <label className="checkbox-row"><input type="checkbox" checked={tocSettings.dotLeaders} onChange={(event) => commitNavigation({ tocSettings: { ...tocSettings, dotLeaders: event.currentTarget.checked } })} />Dot leaders</label>
                  <label className="checkbox-row"><input type="checkbox" checked={tocSettings.includePageLabels} onChange={(event) => commitNavigation({ tocSettings: { ...tocSettings, includePageLabels: event.currentTarget.checked } })} />Prefer page labels</label>
                  <label className="checkbox-row"><input type="checkbox" checked={tocSettings.includeTocInNumbering} onChange={(event) => commitNavigation({ tocSettings: { ...tocSettings, includeTocInNumbering: event.currentTarget.checked } })} />Include TOC pages in numbering</label>
                  <div className="project-quick-actions">
                    <button className="button primary" onClick={insertOrRegenerateToc} disabled={isBusy || tocEntries.length === 0}>{generatedToc?.pageIds.length ? "Regenerate TOC" : "Insert TOC"}</button>
                    <button className="button ghost" onClick={removeGeneratedToc} disabled={isBusy || !generatedToc?.pageIds.length}>Remove TOC</button>
                  </div>
                  <button className="button ghost full-width" onClick={addManualTocEntry} disabled={isBusy}>Add manual entry for current page</button>
                  {manualTocEntries.map((entry) => (
                    <div className="asset-row" key={entry.id}>
                      <span className="asset-type">manual toc</span>
                      <input value={entry.title} onChange={(event) => updateManualTocEntry(entry.id, { title: event.currentTarget.value })} aria-label="Manual TOC title" />
                      <small>{formatBookmarkDestination({ pageId: entry.pageId } as DocumentBookmark, pages)}</small>
                    </div>
                  ))}
                </details>
                <details className="publishing-section">
                  <summary>Navigation Preview</summary>
                  {generatedToc?.stale ? <p className="status-message warning compact">TOC is out of date. Regenerate before export.</p> : null}
                  {navigationIssues.length > 0 ? navigationIssues.map((issue) => <p className="field-error" key={issue.id}>{issue.message}</p>) : <p className="muted-text">Navigation links are valid.</p>}
                  <p className="muted-text">{bookmarks.length} bookmarks · {tocEntries.length} TOC entries · {generatedToc?.pageIds.length ?? 0} generated TOC pages</p>
                </details>
                <div className="project-assets-panel">
                  <div className="section-heading">
                    <strong>Project assets</strong>
                    <span>{projectAssets.length} items</span>
                  </div>
                  {projectAssets.length > 0 ? (
                    projectAssets.map((asset) => (
                      <div className="asset-row" key={asset.id}>
                        <span className="asset-type">{asset.type.replace("-", " ")}</span>
                        <strong>{asset.name}</strong>
                        <small>{formatAssetSize(asset.size)} · {asset.usageCount} used</small>
                        <div className="asset-actions">
                          <button onClick={() => locateProjectAsset(asset.id)}>Locate</button>
                          <button onClick={() => renameProjectAsset(asset.id)}>Rename</button>
                          <button onClick={() => removeUnusedAsset(asset.id)} disabled={asset.usageCount > 0}>Remove unused</button>
                        </div>
                      </div>
                    ))
                  ) : (
                    <p className="muted-text">No embedded project assets yet.</p>
                  )}
                </div>
              </div>
            ) : null}

            {activeWorkspace === "publish" ? (
              <div className="workspace-panel-content publish-summary">
                <div className="publish-ready">
                  <Download size={24} />
                  <strong>Output Preview</strong>
                  <span>{pages.length} pages · {marks.length} edits · {comments.length} comments · publishing marks previewed live</span>
                </div>
                {publishingValidationMessage ? <p className="status-message error compact" role="alert">{publishingValidationMessage}</p> : null}
                <details className="publishing-section" open>
                  <summary>Page Numbers</summary>
                  <label className="checkbox-row">
                    <input
                      type="checkbox"
                      checked={publishingSettings.pageNumbers.enabled}
                      onChange={(event) => {
                        const enabled = event.currentTarget.checked;
                        updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, enabled } }));
                      }}
                    />
                    Enable page numbers
                  </label>
                  <label>Position<select value={publishingSettings.pageNumbers.zone} onChange={(event) => { const zone = event.currentTarget.value as PublishingZone; updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, zone } })); }}>{publishingZoneOptions.map((zone) => <option key={zone.value} value={zone.value}>{zone.label}</option>)}</select></label>
                  <label>Format<select value={publishingSettings.pageNumbers.format} onChange={(event) => { const format = event.currentTarget.value as PublishingNumberFormat; updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, format } })); }}>{pageNumberFormatOptions.map((format) => <option key={format.value} value={format.value}>{format.label}</option>)}</select></label>
                  {publishingSettings.pageNumbers.format === "custom" ? <label>Template<input value={publishingSettings.pageNumbers.customTemplate} onChange={(event) => { const customTemplate = event.currentTarget.value; updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, customTemplate } }), false); }} placeholder="Page {page} of {pages}" /></label> : null}
                  <div className="page-field-grid">
                    <label>Start<input type="number" value={publishingSettings.pageNumbers.startNumber} onChange={(event) => { const startNumber = Number(event.currentTarget.value) || 1; updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, startNumber } })); }} /></label>
                    <label>Offset<input type="number" value={publishingSettings.pageNumbers.offset} onChange={(event) => { const offset = Number(event.currentTarget.value) || 0; updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, offset } })); }} /></label>
                  </div>
                  <div className="page-field-grid">
                    <label>Prefix<input value={publishingSettings.pageNumbers.prefix} onChange={(event) => { const prefix = event.currentTarget.value; updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, prefix } }), false); }} /></label>
                    <label>Suffix<input value={publishingSettings.pageNumbers.suffix} onChange={(event) => { const suffix = event.currentTarget.value; updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, suffix } }), false); }} /></label>
                  </div>
                  <div className="page-field-grid">
                    <label>Distance<input type="number" min="0" value={publishingSettings.pageNumbers.distanceFromEdge} onChange={(event) => { const distanceFromEdge = Math.max(0, Number(event.currentTarget.value) || 0); updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, distanceFromEdge } })); }} /></label>
                    <label>Horizontal offset<input type="number" value={publishingSettings.pageNumbers.horizontalOffset} onChange={(event) => { const horizontalOffset = Number(event.currentTarget.value) || 0; updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, horizontalOffset } })); }} /></label>
                  </div>
                  <label>Units<select value={publishingSettings.pageNumbers.unit} onChange={(event) => { const unit = event.currentTarget.value as PublishingUnit; updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, unit } })); }}><option value="mm">Millimetres</option><option value="cm">Centimetres</option><option value="in">Inches</option><option value="pt">Points</option></select></label>
                  <PublishingTargetControls target={publishingSettings.pageNumbers.target} pageCount={pages.length} onChange={(target) => updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, target } }))} />
                  <details className="publishing-zone-control">
                    <summary>Numbering Sections</summary>
                    <div className="project-quick-actions">
                      <button className="button ghost" type="button" onClick={() => addNumberingSection()} disabled={isBusy}>Add section</button>
                      <button className="button ghost" type="button" onClick={createDefaultNumberingWorkflow} disabled={isBusy}>Cover / TOC / Main / Appendix</button>
                      <button className="button ghost" type="button" onClick={() => {
                        const first = pages[0];
                        if (first) commitNavigation({ numberingSections: [defaultLegacyNumberingSection(first.id, publishingSettings.pageNumbers.format, publishingSettings.pageNumbers.startNumber, publishingSettings.pageNumbers.prefix, publishingSettings.pageNumbers.suffix)] });
                      }} disabled={isBusy}>Use legacy settings</button>
                    </div>
                    {numberingIssues.map((issue) => <p className="field-error" key={issue.id}>{issue.message}</p>)}
                    <div className="navigation-preview-list">
                      {numberingSummary.length > 0 ? numberingSummary.map(({ section, text }) => (
                        <div className="numbering-section-row" key={section.id}>
                          <input value={section.label} onChange={(event) => updateNumberingSection(section.id, { label: event.currentTarget.value })} aria-label="Numbering section label" />
                          <small>{text}</small>
                          <div className="page-field-grid">
                            <label>Start<select value={section.startPageId} onChange={(event) => updateNumberingSection(section.id, { startPageId: event.currentTarget.value })}>{pages.map((page) => <option key={page.id} value={page.id}>{page.label || `Page ${page.pageNumber}`}</option>)}</select></label>
                            <label>End<select value={section.endPageId ?? ""} onChange={(event) => updateNumberingSection(section.id, { endPageId: event.currentTarget.value || undefined })}><option value="">Auto</option>{pages.map((page) => <option key={page.id} value={page.id}>{page.label || `Page ${page.pageNumber}`}</option>)}</select></label>
                          </div>
                          <div className="page-field-grid">
                            <label>Format<select value={section.format} onChange={(event) => updateNumberingSection(section.id, { format: event.currentTarget.value as PublishingNumberFormat })}>{pageNumberFormatOptions.map((format) => <option key={format.value} value={format.value}>{format.label}</option>)}</select></label>
                            <label>Start<input type="number" min="1" value={section.startValue} onChange={(event) => updateNumberingSection(section.id, { startValue: Math.max(1, Number(event.currentTarget.value) || 1) })} /></label>
                          </div>
                          <div className="page-field-grid">
                            <label>Prefix<input value={section.prefix} onChange={(event) => updateNumberingSection(section.id, { prefix: event.currentTarget.value })} /></label>
                            <label>Suffix<input value={section.suffix} onChange={(event) => updateNumberingSection(section.id, { suffix: event.currentTarget.value })} /></label>
                          </div>
                          <label className="checkbox-row"><input type="checkbox" checked={section.includeNumbering} onChange={(event) => updateNumberingSection(section.id, { includeNumbering: event.currentTarget.checked })} />Show numbering</label>
                          <label className="checkbox-row"><input type="checkbox" checked={section.includeInTotal} onChange={(event) => updateNumberingSection(section.id, { includeInTotal: event.currentTarget.checked })} />Include in total</label>
                          <label className="checkbox-row"><input type="checkbox" checked={section.restart} onChange={(event) => updateNumberingSection(section.id, { restart: event.currentTarget.checked })} />Restart numbering</label>
                          <label className="checkbox-row"><input type="checkbox" checked={section.usePageLabel} onChange={(event) => updateNumberingSection(section.id, { usePageLabel: event.currentTarget.checked })} />Display page label</label>
                          <div className="bookmark-row-actions">
                            <button type="button" onClick={() => moveNumberingSection(section.id, -1)}>Up</button>
                            <button type="button" onClick={() => moveNumberingSection(section.id, 1)}>Down</button>
                            <button type="button" onClick={() => duplicateSection(section)}>Duplicate</button>
                            <button type="button" onClick={() => deleteNumberingSection(section.id)}>Delete</button>
                          </div>
                        </div>
                      )) : <p className="muted-text">No numbering sections. Legacy page numbering is still used until you add sections.</p>}
                    </div>
                  </details>
                  <PublishingStyleControls style={publishingSettings.pageNumbers.style} onChange={(style) => updatePublishingSettings((settings) => ({ ...settings, pageNumbers: { ...settings.pageNumbers, style: { ...settings.pageNumbers.style, ...style } } }))} />
                  <button className="button ghost full-width" onClick={() => saveCurrentPublishingPreset("pageNumbers")}>Save page-number preset</button>
                </details>
                <details className="publishing-section">
                  <summary>Headers & Footers</summary>
                  <label className="checkbox-row"><input type="checkbox" checked={publishingSettings.headerFooter.enabled} onChange={(event) => { const enabled = event.currentTarget.checked; updatePublishingSettings((settings) => ({ ...settings, headerFooter: { ...settings.headerFooter, enabled } })); }} />Enable headers and footers</label>
                  <label className="checkbox-row"><input type="checkbox" checked={publishingSettings.headerFooter.advanced} onChange={(event) => { const advanced = event.currentTarget.checked; updatePublishingSettings((settings) => ({ ...settings, headerFooter: { ...settings.headerFooter, advanced } })); }} />Advanced odd/even/first page mode</label>
                  <PublishingHeaderFooterZoneControls title="Header left" zone={publishingSettings.headerFooter.header.left} imageAssets={publishingImageAssets} onTextChange={(text) => updateHeaderFooterZone("header", "left", { text })} onImageChange={(patch) => updateHeaderFooterZoneImage("header", "left", patch)} onRemoveImage={() => updateHeaderFooterZone("header", "left", { image: undefined })} onImportImage={() => startProjectImageAssetImport({ kind: "headerFooter", area: "header", side: "left" })} />
                  <PublishingHeaderFooterZoneControls title="Header centre" zone={publishingSettings.headerFooter.header.center} imageAssets={publishingImageAssets} onTextChange={(text) => updateHeaderFooterZone("header", "center", { text })} onImageChange={(patch) => updateHeaderFooterZoneImage("header", "center", patch)} onRemoveImage={() => updateHeaderFooterZone("header", "center", { image: undefined })} onImportImage={() => startProjectImageAssetImport({ kind: "headerFooter", area: "header", side: "center" })} />
                  <PublishingHeaderFooterZoneControls title="Header right" zone={publishingSettings.headerFooter.header.right} imageAssets={publishingImageAssets} onTextChange={(text) => updateHeaderFooterZone("header", "right", { text })} onImageChange={(patch) => updateHeaderFooterZoneImage("header", "right", patch)} onRemoveImage={() => updateHeaderFooterZone("header", "right", { image: undefined })} onImportImage={() => startProjectImageAssetImport({ kind: "headerFooter", area: "header", side: "right" })} />
                  <PublishingHeaderFooterZoneControls title="Footer left" zone={publishingSettings.headerFooter.footer.left} imageAssets={publishingImageAssets} onTextChange={(text) => updateHeaderFooterZone("footer", "left", { text })} onImageChange={(patch) => updateHeaderFooterZoneImage("footer", "left", patch)} onRemoveImage={() => updateHeaderFooterZone("footer", "left", { image: undefined })} onImportImage={() => startProjectImageAssetImport({ kind: "headerFooter", area: "footer", side: "left" })} />
                  <PublishingHeaderFooterZoneControls title="Footer centre" zone={publishingSettings.headerFooter.footer.center} imageAssets={publishingImageAssets} onTextChange={(text) => updateHeaderFooterZone("footer", "center", { text })} onImageChange={(patch) => updateHeaderFooterZoneImage("footer", "center", patch)} onRemoveImage={() => updateHeaderFooterZone("footer", "center", { image: undefined })} onImportImage={() => startProjectImageAssetImport({ kind: "headerFooter", area: "footer", side: "center" })} />
                  <PublishingHeaderFooterZoneControls title="Footer right" zone={publishingSettings.headerFooter.footer.right} imageAssets={publishingImageAssets} onTextChange={(text) => updateHeaderFooterZone("footer", "right", { text })} onImageChange={(patch) => updateHeaderFooterZoneImage("footer", "right", patch)} onRemoveImage={() => updateHeaderFooterZone("footer", "right", { image: undefined })} onImportImage={() => startProjectImageAssetImport({ kind: "headerFooter", area: "footer", side: "right" })} />
                  {publishingSettings.headerFooter.advanced ? (
                    <PublishingAdvancedHeaderFooterControls
                      settings={publishingSettings}
                      imageAssets={publishingImageAssets}
                      onTextChange={(key, zone, text) => updateHeaderFooterOverrideZone(key, zone, { text })}
                      onImageChange={updateHeaderFooterOverrideImage}
                      onRemoveImage={(key, zone) => updateHeaderFooterOverrideZone(key, zone, { image: undefined })}
                      onImportImage={(key, zone) => startProjectImageAssetImport({ kind: "headerFooterOverride", key, zone })}
                    />
                  ) : null}
                  <div className="page-field-grid">
                    <label>Margin<input type="number" min="0" value={publishingSettings.headerFooter.margin} onChange={(event) => { const margin = Math.max(0, Number(event.currentTarget.value) || 0); updatePublishingSettings((settings) => ({ ...settings, headerFooter: { ...settings.headerFooter, margin } })); }} /></label>
                    <label>Units<select value={publishingSettings.headerFooter.unit} onChange={(event) => { const unit = event.currentTarget.value as PublishingUnit; updatePublishingSettings((settings) => ({ ...settings, headerFooter: { ...settings.headerFooter, unit } })); }}><option value="mm">Millimetres</option><option value="cm">Centimetres</option><option value="in">Inches</option><option value="pt">Points</option></select></label>
                  </div>
                  <PublishingTargetControls target={publishingSettings.headerFooter.target} pageCount={pages.length} onChange={(target) => updatePublishingSettings((settings) => ({ ...settings, headerFooter: { ...settings.headerFooter, target } }))} />
                  <PublishingStyleControls style={publishingSettings.headerFooter.style} onChange={(style) => updatePublishingSettings((settings) => ({ ...settings, headerFooter: { ...settings.headerFooter, style: { ...settings.headerFooter.style, ...style } } }))} />
                  <button className="button ghost full-width" onClick={() => saveCurrentPublishingPreset("headerFooter")}>Save header/footer preset</button>
                </details>
                <details className="publishing-section">
                  <summary>Watermark</summary>
                  <label className="checkbox-row"><input type="checkbox" checked={publishingSettings.watermark.enabled} onChange={(event) => { const enabled = event.currentTarget.checked; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, enabled } })); }} />Enable watermark</label>
                  <label>Type<select value={publishingSettings.watermark.type} onChange={(event) => { const type = event.currentTarget.value as "text" | "image"; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, type } })); }}><option value="text">Text</option><option value="image">Image</option></select></label>
                  {publishingSettings.watermark.type === "text" ? (
                    <>
                      <label>Preset<select value={publishingSettings.watermark.preset} onChange={(event) => { const preset = event.currentTarget.value; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, preset, text: preset } })); }}>{publishingWatermarkPresets.map((preset) => <option key={preset} value={preset}>{preset}</option>)}</select></label>
                      <label>Text<input value={publishingSettings.watermark.text} onChange={(event) => { const text = event.currentTarget.value; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, text } }), false); }} /></label>
                    </>
                  ) : (
                    <>
                      <label>Image asset<select value={publishingSettings.watermark.imageAssetId ?? ""} onChange={(event) => { const imageAssetId = event.currentTarget.value || undefined; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, imageAssetId } })); }}><option value="">Choose image asset</option>{publishingImageAssets.map((asset) => <option key={asset.id} value={asset.id}>{asset.name}</option>)}</select></label>
                      <button className="button ghost full-width" type="button" onClick={() => startProjectImageAssetImport({ kind: "watermark" })}>Import image asset</button>
                    </>
                  )}
                  <label>Position<select value={publishingSettings.watermark.position} onChange={(event) => { const position = event.currentTarget.value as PublishingPositionPreset; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, position } })); }}>{watermarkPositionOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}</select></label>
                  {publishingSettings.watermark.position === "custom" ? (
                    <div className="page-field-grid">
                      <label>X<input type="number" value={publishingSettings.watermark.customX} onChange={(event) => { const customX = Number(event.currentTarget.value) || 0; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, customX } })); }} /></label>
                      <label>Y<input type="number" value={publishingSettings.watermark.customY} onChange={(event) => { const customY = Number(event.currentTarget.value) || 0; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, customY } })); }} /></label>
                    </div>
                  ) : null}
                  <div className="page-field-grid">
                    <label>Opacity<input type="number" min="0.01" max="1" step="0.05" value={publishingSettings.watermark.opacity} onChange={(event) => { const opacity = clamp(Number(event.currentTarget.value), 0.01, 1); updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, opacity } })); }} /></label>
                    <label>Rotation<input type="number" value={publishingSettings.watermark.rotation} onChange={(event) => { const rotation = Number(event.currentTarget.value) || 0; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, rotation } })); }} /></label>
                  </div>
                  <label>Scale<input type="range" min="0.25" max="3" step="0.05" value={publishingSettings.watermark.scale} onChange={(event) => { const scale = Number(event.currentTarget.value); updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, scale } })); }} /></label>
                  <div className="page-field-grid">
                    <label>Units<select value={publishingSettings.watermark.unit} onChange={(event) => { const unit = event.currentTarget.value as PublishingUnit; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, unit } })); }}><option value="mm">Millimetres</option><option value="cm">Centimetres</option><option value="in">Inches</option><option value="pt">Points</option></select></label>
                    <label className="checkbox-row"><input type="checkbox" checked={publishingSettings.watermark.tiled} onChange={(event) => { const tiled = event.currentTarget.checked; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, tiled } })); }} />Tile watermark</label>
                  </div>
                  <label className="checkbox-row"><input type="checkbox" checked={publishingSettings.watermark.maintainAspectRatio} onChange={(event) => { const maintainAspectRatio = event.currentTarget.checked; updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, maintainAspectRatio } })); }} />Maintain image aspect ratio</label>
                  <PublishingTargetControls target={publishingSettings.watermark.target} pageCount={pages.length} onChange={(target) => updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, target } }))} />
                  <PublishingStyleControls style={publishingSettings.watermark.style} onChange={(style) => updatePublishingSettings((settings) => ({ ...settings, watermark: { ...settings.watermark, style: { ...settings.watermark.style, ...style } } }))} />
                  <button className="button ghost full-width" onClick={() => saveCurrentPublishingPreset("watermark")}>Save watermark preset</button>
                </details>
                {publishingPresets.length > 0 ? (
                  <details className="publishing-section">
                    <summary>Local Presets</summary>
                    {publishingPresets.map((preset) => (
                      <div className="preset-row" key={preset.id}>
                        <button onClick={() => applyPublishingPreset(preset)}>{preset.name}<span>{preset.type}</span></button>
                        <button className="link-button" onClick={() => renamePublishingPreset(preset)}>Rename</button>
                        <button className="link-button" onClick={() => { deletePublishingPreset(preset.id); setPublishingPresets(loadPublishingPresets()); }}>Delete</button>
                      </div>
                    ))}
                  </details>
                ) : null}
                <button className="button primary full-width" onClick={() => void exportPdf()} disabled={isBusy} title="Export edited PDF" aria-label="Export edited PDF">
                  Export PDF
                </button>
                <button className="button ghost full-width" onClick={() => setActivePageDialog("extract")} disabled={isBusy} title="Extract selected pages" aria-label="Extract selected pages">
                  Export selected pages
                </button>
              </div>
            ) : null}

            {leftPanel === "bookmarks" ? (
              <BookmarkTreePanel
                bookmarks={visibleBookmarks}
                pages={pages}
                selectedBookmarkId={selectedBookmarkId}
                issues={navigationIssues}
                onSelect={(bookmark) => navigateToBookmark(bookmark)}
                onToggle={(bookmark) => updateBookmark(bookmark.id, { expanded: !bookmark.expanded })}
                onRename={(bookmark, title) => updateBookmark(bookmark.id, { title })}
                onMove={moveSelectedBookmark}
                onNest={nestSelectedBookmark}
                onOutdent={outdentSelectedBookmark}
                onDuplicate={duplicateBookmark}
                onDelete={removeBookmark}
                onRepair={repairBookmarkDestination}
              />
            ) : null}

            {activeWorkspace === "review" ? (
              <div className="panel-list">
                {comments.length > 0 ? (
                  comments.map((mark) => (
                    <button key={mark.id} className="left-comment" onClick={() => selectComment(mark)} title={getMarkLabel(mark)} aria-label={getMarkLabel(mark)}>
                      <span>{mark.comment?.resolved ? "Resolved" : "Open"}</span>
                      <strong>{getCommentPreview(mark.comment?.body ?? "")}</strong>
                      <small>Page {mark.page}</small>
                    </button>
                  ))
                ) : (
                  <div className="panel-empty">
                    <MessageSquare size={22} />
                    <strong>No comments</strong>
                    <span>Add a comment from the toolbar.</span>
                  </div>
                )}
              </div>
            ) : null}

            {leftPanel === "search" && activeWorkspace !== "review" ? (
              <div className="panel-empty">
                <Search size={22} />
                <strong>Page search</strong>
                <div className="searchbox full">
                  <Search size={15} />
                  <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Find page number" aria-label="Find page number" />
                </div>
              </div>
            ) : null}
          </aside>
        ) : null}

        <section className="editor">

          {errorMessage ? (
            <div className="status-message error" role="alert">
              {errorMessage}
            </div>
          ) : null}
          {!errorMessage && projectStatusMessage ? (
            <div className="status-message project" role="status" aria-live="polite">
              {projectStatusMessage}
            </div>
          ) : null}

          <div
            className={`canvas-stage ${isDragging ? "dragging" : ""}`}
            onDragOver={(event) => {
              event.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(event) => {
              event.preventDefault();
              setIsDragging(false);
              const files = Array.from(event.dataTransfer.files);
              const projectFile = files.find((file) => file.name.toLowerCase().endsWith(".pproj"));
              if (projectFile) {
                void openProjectFile(projectFile);
                return;
              }
              const docxFile = files.find((file) => file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || file.name.toLowerCase().endsWith(".docx") || file.name.toLowerCase().endsWith(".doc"));
              if (docxFile) {
                beginDocxImport(docxFile, hasOpenProject ? "append" : "replace");
                return;
              }
              const pptxFile = files.find((file) => file.type === "application/vnd.openxmlformats-officedocument.presentationml.presentation" || file.name.toLowerCase().endsWith(".pptx") || file.name.toLowerCase().endsWith(".ppt"));
              if (pptxFile) {
                beginPptxImport(pptxFile, hasOpenProject ? "append" : "replace");
                return;
              }
              const pdfFiles = files.filter((file) => file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf"));
              const file = pdfFiles[0];
              if (!file) return;
              if (pdfBytes || pages.some((page) => page.sourceDocumentId)) {
                void importPdfFiles(pdfFiles, "after");
                return;
              }
              void loadFile(file);
            }}
          >
            {workState ? <ProgressOverlay message={workState.message} progress={workState.progress} /> : null}
            {!hasOpenProject ? (
              <WelcomeScreen
                logoSrc={BRAND_ICON_SRC}
                recentProjects={recentProjects}
                autosave={autosaveCandidate}
                onNewProject={startNewProject}
                onOpenProject={() => void chooseProjectFile()}
                onOpenPdf={() => void choosePdfFile()}
                onOpenDocx={() => void chooseDocxFile()}
                onOpenPptx={() => void choosePptxFile()}
                onOpenRecent={reopenRecentProject}
                onRemoveRecent={(id) => {
                  void (async () => {
                    if (runtime.isDesktop) {
                      await runtime.removeRecentProject(id);
                    } else {
                      removeBrowserRecentProject(id);
                    }
                    await refreshStoredRecentProjects();
                  })();
                }}
                onClearRecent={() => {
                  void (async () => {
                    if (runtime.isDesktop) {
                      await runtime.clearRecentProjects();
                    } else {
                      clearBrowserRecentProjects();
                    }
                    setRecentProjects([]);
                  })();
                }}
                onRestoreAutosave={restoreAutosave}
              />
            ) : (
            pages
              .filter((page) => page.pageNumber === currentPage)
              .map((page) => (
                <DocumentPage
                  key={page.pageNumber}
                  page={page}
                  marks={marks.filter((mark) => markBelongsToPage(mark, page))}
                  zoom={zoom}
                  selectedMark={selectedMark}
                  activeTool={activeTool}
                  penColor={penColor}
                  penWidth={penWidth}
                  penOpacity={penOpacity}
                  pendingSignature={pendingSignature}
                  publishingPreviewItems={getPublishingPreviewItems({
                    settings: publishingSettings,
                    page,
                    pages,
                    currentPage,
                    selectedPageIds,
                    tokenContext: getPublishingTokenContext(),
                    imageAssetIds: publishingImageAssets.map((asset) => asset.id),
                    resolvedPageNumbers,
                  })}
                  publishingImageAssets={publishingImageAssets}
                  onSelect={setSelectedMark}
                  onAddMark={addMark}
                  onAddTextBox={addTextBox}
                  onAddShape={addShape}
                  onAddStroke={addStroke}
                  onAddTextMarkup={addTextMarkup}
                  onPlaceSignature={placeSignature}
                  onPreviewMark={previewMark}
                  onCommitMarkChange={commitMarkChange}
                  editingText={editingText}
                  onStartTextEdit={startTextEdit}
                  onTextDraftChange={(draft) => setEditingText((existing) => (existing ? { ...existing, draft } : existing))}
                  onSaveTextEdit={saveTextEdit}
                  onCancelTextEdit={cancelTextEdit}
                />
              ))
            )}
          </div>
        </section>

        {!isRightPanelCollapsed ? (
        <aside className="inspector">
          <div className="inspector-header">
            <div>
              <h2>{getWorkspaceInspectorTitle(activeWorkspace)}</h2>
              <span>{canShowObjectInspector && selected ? getMarkLabel(selected) : getWorkspaceInspectorSubtitle(activeWorkspace)}</span>
            </div>
            <button className="panel-icon-button" onClick={() => setIsRightPanelCollapsed(true)} title="Collapse right panel" aria-label="Collapse right panel">
              <PanelRightClose size={16} />
            </button>
          </div>
          <details className="inspector-section" open>
            <summary>{canShowObjectInspector && selected ? "Selection" : getWorkspaceInspectorTitle(activeWorkspace)}</summary>
            <div className="inspector-section-body">
          {activeWorkspace === "assemble" ? (
            <div className="control-stack workspace-properties">
              <div className="property-readout">
                <span>Current page</span>
                <strong>{currentPage}</strong>
              </div>
              <div className="property-readout">
                <span>Size</span>
                <strong>{currentPageView ? `${Math.round(currentPageView.width / PAGE_SCALE)} x ${Math.round(currentPageView.height / PAGE_SCALE)} pt` : "No page"}</strong>
              </div>
              <div className="property-readout">
                <span>Orientation</span>
                <strong>{currentPageView && currentPageView.width >= currentPageView.height ? "Landscape" : "Portrait"}</strong>
              </div>
              <div className="property-readout">
                <span>Rotation</span>
                <strong>{currentPageView?.rotation ?? 0}deg</strong>
              </div>
              {currentPageImportMetadata ? (
                <div className="docx-import-status">
                  <strong>{isPptxImportMetadata(currentPageImportMetadata) ? "Imported from PowerPoint" : "Imported from Word"}</strong>
                  <span>{currentPageImportMetadata.sourceName} · {currentPageImportMetadata.fidelityMode} · {currentPageView?.sourceMappings?.length ?? 0} mapped blocks</span>
                  {isPptxImportMetadata(currentPageImportMetadata) ? (
                    <button className="button ghost full-width" onClick={() => beginPptxReimport(currentPageImportMetadata)} disabled={isBusy || !currentPageImportMetadata.originalSourceDocumentId}>Re-import source</button>
                  ) : (
                    <button className="button ghost full-width" onClick={() => beginDocxReimport(currentPageImportMetadata)} disabled={isBusy || !currentPageImportMetadata.originalSourceDocumentId}>Re-import source</button>
                  )}
                </div>
              ) : null}
              <label>
                Selected page label
                <input
                  value={pages.find((page) => selectedPageIds.includes(page.id))?.label ?? ""}
                  onChange={(event) => updateSelectedPageLabel(event.target.value)}
                  placeholder="Cover, Section A, A-01..."
                  disabled={isBusy}
                />
              </label>
            </div>
          ) : null}
          {activeWorkspace === "import" ? (
            <div className="control-stack workspace-properties">
              <div className="property-readout">
                <span>Project</span>
                <strong>{projectMetadata.name}</strong>
              </div>
              <div className="property-readout">
                <span>Pages</span>
                <strong>{pages.length}</strong>
              </div>
              <div className="property-readout">
                <span>Source files</span>
                <strong>{Object.keys(sourceDocuments).length || 1}</strong>
              </div>
              <div className="property-readout">
                <span>Format</span>
                <strong>.pproj v1</strong>
              </div>
              {docxImportGroups.length > 0 ? (
                <details className="publishing-section" open>
                  <summary>Word imports</summary>
                  <div className="docx-import-list">
                    {docxImportGroups.map((group) => (
                      <div className="docx-import-row" key={group.metadata.importId}>
                        <strong>{group.metadata.sourceName}</strong>
                        <span>{group.pages.length} pages · {group.metadata.fidelityMode} · {group.mappings} mappings · {group.warnings} warnings</span>
                        <div className="dialog-actions compact">
                          <button className="button ghost" onClick={() => { setCurrentPage(group.pages[0]?.pageNumber ?? 1); setSelectedPageIds(group.pages[0] ? [group.pages[0].id] : []); }} disabled={!group.pages[0]}>Review</button>
                          <button className="button ghost" onClick={() => beginDocxReimport(group.metadata)} disabled={isBusy || !group.metadata.originalSourceDocumentId}>Re-import</button>
                        </div>
                      </div>
                    ))}
                  </div>
                </details>
              ) : null}
              {pptxImportGroups.length > 0 ? (
                <details className="publishing-section" open>
                  <summary>PowerPoint imports</summary>
                  <div className="docx-import-list">
                    {pptxImportGroups.map((group) => (
                      <div className="docx-import-row" key={group.metadata.importId}>
                        <strong>{group.metadata.sourceName}</strong>
                        <span>{group.pages.length} slides · {group.metadata.fidelityMode} · {group.mappings} mappings · {group.notes} notes · {group.warnings} warnings</span>
                        <div className="dialog-actions compact">
                          <button className="button ghost" onClick={() => { setCurrentPage(group.pages[0]?.pageNumber ?? 1); setSelectedPageIds(group.pages[0] ? [group.pages[0].id] : []); }} disabled={!group.pages[0]}>Review</button>
                          <button className="button ghost" onClick={() => beginPptxReimport(group.metadata)} disabled={isBusy || !group.metadata.originalSourceDocumentId}>Re-import</button>
                        </div>
                      </div>
                    ))}
                  </div>
                </details>
              ) : null}
              <div className="property-readout">
                <span>Last saved</span>
                <strong>{lastSavedAt ? formatDateTime(lastSavedAt) : "Not saved"}</strong>
              </div>
              <div className="property-readout">
                <span>Autosave</span>
                <strong>{lastAutosavedAt ? formatDateTime(lastAutosavedAt) : "Pending"}</strong>
              </div>
              <label>
                Project name
                <input value={projectMetadata.name} onChange={(event) => updateProjectMetadata({ name: event.currentTarget.value })} />
              </label>
              <label>
                Client
                <input value={projectMetadata.client} onChange={(event) => updateProjectMetadata({ client: event.currentTarget.value })} />
              </label>
              <label>
                Author
                <input value={projectMetadata.author} onChange={(event) => updateProjectMetadata({ author: event.currentTarget.value })} />
              </label>
              <label>
                Description
                <textarea value={projectMetadata.description} onChange={(event) => updateProjectMetadata({ description: event.currentTarget.value })} rows={3} />
              </label>
              <label>
                Tags
                <input value={metadataTagsToInput(projectMetadata)} onChange={(event) => updateProjectMetadata({ tags: inputToMetadataTags(event.currentTarget.value) })} placeholder="bid, review, client" />
              </label>
              <div className="property-readout">
                <span>Created</span>
                <strong>{formatDateTime(projectMetadata.createdAt)}</strong>
              </div>
              <div className="property-readout">
                <span>Modified</span>
                <strong>{formatDateTime(projectMetadata.modifiedAt)}</strong>
              </div>
              <button className="button ghost full-width" onClick={() => setActivePageDialog("merge")} disabled={isBusy} title="Combine PDF files" aria-label="Combine PDF files">
                Combine PDFs
              </button>
              <button className="button primary full-width" onClick={() => void saveProject()} disabled={isBusy || !hasOpenProject} title="Save Publish Pro project" aria-label="Save Publish Pro project">
                Save Project
              </button>
            </div>
          ) : null}
          {activeWorkspace === "publish" ? (
            <div className="control-stack workspace-properties">
              <div className="property-readout">
                <span>Output</span>
                <strong>PDF</strong>
              </div>
              <div className="property-readout">
                <span>Pages</span>
                <strong>{pages.length}</strong>
              </div>
              <div className="property-readout">
                <span>Edits</span>
                <strong>{marks.length}</strong>
              </div>
              <div className="property-readout">
                <span>Comments</span>
                <strong>{comments.length}</strong>
              </div>
              <button className="button primary full-width" onClick={() => void exportPdf()} disabled={isBusy} title="Export edited PDF" aria-label="Export edited PDF">
                Export PDF
              </button>
              {runtime.isDesktop && lastExportPath ? (
                <div className="dialog-actions compact">
                  <button className="button ghost" onClick={() => void runtime.openPath(lastExportPath)} disabled={isBusy}>Open PDF</button>
                  <button className="button ghost" onClick={() => void runtime.revealPath(lastExportPath)} disabled={isBusy}>Show in folder</button>
                </div>
              ) : null}
            </div>
          ) : null}
          {activeWorkspace === "review" && activeTool === "draw" ? (
            <div className="control-stack pen-controls">
              <label>
                Stroke colour
                <input type="color" value={penColor} onChange={(event) => setPenColor(event.currentTarget.value)} aria-label="Pen stroke colour" />
              </label>
              <label>
                Colour hex
                <input
                  type="text"
                  value={penColor}
                  onChange={(event) => {
                    const color = normalizeHexColor(event.currentTarget.value);
                    if (color) setPenColor(color);
                  }}
                  aria-label="Pen stroke colour hex"
                />
              </label>
              <label>
                Stroke width
                <input
                  type="range"
                  min="1"
                  max="24"
                  value={penWidth}
                  onChange={(event) => setPenWidth(Number(event.currentTarget.value))}
                  aria-label="Pen stroke width"
                />
              </label>
              <label>
                Width value
                <input
                  type="number"
                  min="1"
                  max="24"
                  value={penWidth}
                  onChange={(event) => setPenWidth(clamp(Number(event.currentTarget.value), 1, 24))}
                  aria-label="Pen stroke width value"
                />
              </label>
              <label>
                Opacity
                <input
                  type="range"
                  min="0.1"
                  max="1"
                  step="0.05"
                  value={penOpacity}
                  onChange={(event) => setPenOpacity(Number(event.currentTarget.value))}
                  aria-label="Pen opacity"
                />
              </label>
              <label>
                Opacity value
                <input
                  type="number"
                  min="0.1"
                  max="1"
                  step="0.05"
                  value={penOpacity}
                  onChange={(event) => setPenOpacity(clamp(Number(event.currentTarget.value), 0.1, 1))}
                  aria-label="Pen opacity value"
                />
              </label>
            </div>
          ) : null}
          {activeWorkspace === "review" && selected?.kind === "comment" && selected.comment ? (
            <div className="control-stack comment-editor">
              <div className={`comment-status ${selected.comment.resolved ? "resolved" : ""}`}>
                {selected.comment.resolved ? <CheckCircle2 size={16} aria-hidden="true" /> : <MessageSquare size={16} aria-hidden="true" />}
                <span>{selected.comment.resolved ? "Resolved comment" : "Open comment"}</span>
              </div>
              <label>
                Comment
                <textarea
                  ref={commentEditorRef}
                  value={commentDraft}
                  onChange={(event) => setCommentDraft(event.currentTarget.value)}
                  onKeyDown={(event) => {
                    if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
                      event.preventDefault();
                      saveCommentBody();
                    }
                    if (event.key === "Escape") {
                      event.preventDefault();
                      cancelCommentEdit();
                    }
                  }}
                  aria-label="Comment text"
                />
              </label>
              <div className="inspector-actions">
                <button className="button ghost" onClick={saveCommentBody} title="Save comment text" aria-label="Save comment text">
                  Save
                </button>
                <button className="button ghost" onClick={cancelCommentEdit} title="Cancel comment text changes" aria-label="Cancel comment text changes">
                  Cancel
                </button>
              </div>
              <label>
                Marker colour
                <input
                  type="color"
                  value={selected.comment.color}
                  onChange={(event) => updateSelectedComment({ color: event.currentTarget.value })}
                  aria-label="Comment marker colour"
                />
              </label>
              <label>
                Colour hex
                <input
                  type="text"
                  value={selected.comment.color}
                  onChange={(event) => {
                    const color = normalizeHexColor(event.currentTarget.value);
                    if (color) updateSelectedComment({ color });
                  }}
                  aria-label="Comment marker colour hex"
                />
              </label>
              <div className="inspector-actions">
                <button
                  className="button ghost"
                  onClick={() => updateSelectedComment({ resolved: !selected.comment?.resolved })}
                  title={selected.comment.resolved ? "Reopen comment" : "Resolve comment"}
                  aria-label={selected.comment.resolved ? "Reopen comment" : "Resolve comment"}
                >
                  {selected.comment.resolved ? "Reopen" : "Resolve"}
                </button>
                <button className="button ghost" onClick={duplicateSelectedMark} title="Duplicate comment" aria-label="Duplicate comment">
                  Duplicate
                </button>
              </div>
              <div className="reply-box">
                <h3>Replies</h3>
                {selected.comment.replies.length > 0 ? (
                  <div className="reply-list">
                    {selected.comment.replies.map((reply) => (
                      <div className="reply-item" key={reply.id}>
                        <div className="comment-meta">
                          <strong>{reply.author}</strong>
                          <span>{formatCommentDate(reply.updatedAt)}</span>
                        </div>
                        {editingReply?.id === reply.id ? (
                          <>
                            <textarea
                              value={editingReply.body}
                              onChange={(event) => setEditingReply({ id: reply.id, body: event.currentTarget.value })}
                              aria-label="Edit reply text"
                            />
                            <div className="inspector-actions">
                              <button className="button ghost" onClick={() => saveReplyEdit(reply)} title="Save reply" aria-label="Save reply">
                                Save
                              </button>
                              <button className="button ghost" onClick={() => setEditingReply(null)} title="Cancel reply edit" aria-label="Cancel reply edit">
                                Cancel
                              </button>
                            </div>
                          </>
                        ) : (
                          <>
                            <p>{reply.body}</p>
                            <div className="reply-actions">
                              <button className="link-button" onClick={() => setEditingReply({ id: reply.id, body: reply.body })} title="Edit reply" aria-label="Edit reply">
                                Edit
                              </button>
                              <button className="link-button" onClick={() => deleteReply(reply.id)} title="Delete reply" aria-label="Delete reply">
                                Delete
                              </button>
                            </div>
                          </>
                        )}
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="muted-text">No replies yet.</p>
                )}
                <label>
                  New reply
                  <textarea value={commentReplyDraft} onChange={(event) => setCommentReplyDraft(event.currentTarget.value)} aria-label="New reply text" />
                </label>
                <button className="button ghost full-width" onClick={addCommentReply} disabled={!commentReplyDraft.trim()} title="Add reply" aria-label="Add reply">
                  Add reply
                </button>
              </div>
              <button className="button ghost full-width" onClick={removeSelectedMark} title="Delete comment" aria-label="Delete comment">
                Delete comment
              </button>
            </div>
          ) : activeWorkspace === "review" && selected?.kind === "text" ? (
            <div className="control-stack text-box-controls">
              <button className="button ghost full-width" onClick={() => startTextEdit(selected)} title="Edit text box content" aria-label="Edit text box content">
                Edit text
              </button>
              <label>
                Preset
                <select value="" onChange={(event) => applyTextPreset(event.currentTarget.value)} aria-label="Text box preset">
                  <option value="" disabled>
                    Choose preset
                  </option>
                  {textBoxPresets.map((preset) => (
                    <option key={preset.value} value={preset.value}>
                      {preset.label}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Font family
                <select
                  value={getTextBoxStyle(selected).fontFamily}
                  onChange={(event) => updateTextStyle(selected.id, { fontFamily: event.currentTarget.value as TextBoxFontFamily })}
                  aria-label="Text box font family"
                >
                  {textBoxFontFamilies.map((font) => (
                    <option key={font} value={font}>
                      {font}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Font size
                <input
                  type="number"
                  min="6"
                  max="96"
                  value={selected.size}
                  onChange={(event) => updateMark(selected.id, { size: clamp(Number(event.currentTarget.value), 6, 96) })}
                  aria-label="Text box font size"
                />
              </label>
              <label>
                Text colour
                <input type="color" value={selected.color} onChange={(event) => updateMark(selected.id, { color: event.currentTarget.value })} aria-label="Text box text colour" />
              </label>
              <label>
                Text colour hex
                <input
                  type="text"
                  value={selected.color}
                  onChange={(event) => {
                    const color = normalizeHexColor(event.currentTarget.value);
                    if (color) updateMark(selected.id, { color });
                  }}
                  aria-label="Text box text colour hex"
                />
              </label>
              <div className="toggle-grid" role="group" aria-label="Text style toggles">
                <label className="checkbox-row">
                  <input
                    type="checkbox"
                    checked={getTextBoxStyle(selected).bold}
                    onChange={(event) => updateTextStyle(selected.id, { bold: event.currentTarget.checked })}
                    aria-label="Bold text box text"
                  />
                  Bold
                </label>
                <label className="checkbox-row">
                  <input
                    type="checkbox"
                    checked={getTextBoxStyle(selected).italic}
                    onChange={(event) => updateTextStyle(selected.id, { italic: event.currentTarget.checked })}
                    aria-label="Italic text box text"
                  />
                  Italic
                </label>
                <label className="checkbox-row">
                  <input
                    type="checkbox"
                    checked={getTextBoxStyle(selected).underline}
                    onChange={(event) => updateTextStyle(selected.id, { underline: event.currentTarget.checked })}
                    aria-label="Underline text box text"
                  />
                  Underline
                </label>
              </div>
              <div className="segmented" role="group" aria-label="Text alignment">
                {textBoxAlignments.map((align) => (
                  <button
                    className={getTextBoxStyle(selected).align === align ? "active" : ""}
                    key={align}
                    onClick={() => updateTextStyle(selected.id, { align })}
                    title={`Align text ${align}`}
                    aria-label={`Align text ${align}`}
                    aria-pressed={getTextBoxStyle(selected).align === align}
                  >
                    {align}
                  </button>
                ))}
              </div>
              <div className="dimension-grid">
                <label>
                  Line spacing
                  <input
                    type="number"
                    min="0.8"
                    max="3"
                    step="0.05"
                    value={getTextBoxStyle(selected).lineHeight}
                    onChange={(event) => updateTextStyle(selected.id, { lineHeight: clamp(Number(event.currentTarget.value), 0.8, 3) })}
                    aria-label="Text box line spacing"
                  />
                </label>
                <label>
                  Letter spacing
                  <input
                    type="number"
                    min="-2"
                    max="12"
                    step="0.25"
                    value={getTextBoxStyle(selected).letterSpacing}
                    onChange={(event) => updateTextStyle(selected.id, { letterSpacing: clamp(Number(event.currentTarget.value), -2, 12) })}
                    aria-label="Text box letter spacing"
                  />
                </label>
              </div>
              <label>
                Opacity
                <input
                  type="number"
                  min="0.1"
                  max="1"
                  step="0.05"
                  value={selected.opacity ?? 1}
                  onChange={(event) => updateMark(selected.id, { opacity: clamp(Number(event.currentTarget.value), 0.1, 1) })}
                  aria-label="Text box opacity"
                />
              </label>
              <label>
                Background colour
                <input
                  type="color"
                  value={getTextBoxStyle(selected).backgroundColor}
                  onChange={(event) => updateTextStyle(selected.id, { backgroundColor: event.currentTarget.value })}
                  aria-label="Text box background colour"
                />
              </label>
              <label>
                Background opacity
                <input
                  type="number"
                  min="0"
                  max="1"
                  step="0.05"
                  value={getTextBoxStyle(selected).backgroundOpacity}
                  onChange={(event) => updateTextStyle(selected.id, { backgroundOpacity: clamp(Number(event.currentTarget.value), 0, 1) })}
                  aria-label="Text box background opacity"
                />
              </label>
              <label>
                Border colour
                <input
                  type="color"
                  value={getTextBoxStyle(selected).borderColor}
                  onChange={(event) => updateTextStyle(selected.id, { borderColor: event.currentTarget.value })}
                  aria-label="Text box border colour"
                />
              </label>
              <div className="dimension-grid">
                <label>
                  Border width
                  <input
                    type="number"
                    min="0"
                    max="12"
                    step="0.5"
                    value={getTextBoxStyle(selected).borderWidth}
                    onChange={(event) => updateTextStyle(selected.id, { borderWidth: clamp(Number(event.currentTarget.value), 0, 12) })}
                    aria-label="Text box border width"
                  />
                </label>
                <label>
                  Padding
                  <input
                    type="number"
                    min="0"
                    max="48"
                    value={getTextBoxStyle(selected).padding}
                    onChange={(event) => updateTextStyle(selected.id, { padding: clamp(Number(event.currentTarget.value), 0, 48) })}
                    aria-label="Text box internal padding"
                  />
                </label>
              </div>
              <label>
                Rotation
                <input
                  type="number"
                  min="-180"
                  max="180"
                  value={Math.round(selected.rotation)}
                  onChange={(event) => updateMark(selected.id, { rotation: Number(event.currentTarget.value) })}
                  aria-label="Text box rotation degrees"
                />
              </label>
              <div className="dimension-grid">
                <label>
                  W
                  <input
                    type="number"
                    value={Math.round(selected.width)}
                    onChange={(event) => updateMark(selected.id, getDimensionPatch(selected, "width", Number(event.currentTarget.value)))}
                    aria-label="Text box width"
                  />
                </label>
                <label>
                  H
                  <input
                    type="number"
                    value={Math.round(selected.height)}
                    onChange={(event) => updateMark(selected.id, getDimensionPatch(selected, "height", Number(event.currentTarget.value)))}
                    aria-label="Text box height"
                  />
                </label>
              </div>
              <div className="inspector-actions">
                <button className="button ghost" onClick={duplicateSelectedMark} title="Duplicate text box" aria-label="Duplicate text box">
                  Duplicate
                </button>
                <button className="button ghost" onClick={removeSelectedMark} title="Delete text box" aria-label="Delete text box">
                  Delete
                </button>
              </div>
            </div>
          ) : activeWorkspace === "review" && selected?.kind === "shape" && selected.shapeStyle ? (
            <div className="control-stack shape-controls">
              <div className="comment-status">
                <Square size={16} aria-hidden="true" />
                <span>{shapeToolLabels[selected.shapeStyle.shapeKind]}</span>
              </div>
              {selected.shapeStyle.shapeKind === "callout" ? (
                <label>
                  Callout text
                  <textarea
                    value={selected.shapeStyle.calloutText}
                    onChange={(event) => updateShapeStyle(selected.id, { calloutText: event.currentTarget.value })}
                    aria-label="Callout text"
                  />
                </label>
              ) : null}
              <label>
                Stroke colour
                <input type="color" value={selected.shapeStyle.strokeColor} onChange={(event) => updateShapeStyle(selected.id, { strokeColor: event.currentTarget.value })} aria-label="Shape stroke colour" />
              </label>
              <label>
                Fill colour
                <input type="color" value={selected.shapeStyle.fillColor} onChange={(event) => updateShapeStyle(selected.id, { fillColor: event.currentTarget.value })} aria-label="Shape fill colour" />
              </label>
              <div className="dimension-grid">
                <label>
                  Stroke width
                  <input
                    type="number"
                    min="0.5"
                    max="24"
                    step="0.5"
                    value={selected.shapeStyle.strokeWidth}
                    onChange={(event) => updateShapeStyle(selected.id, { strokeWidth: clamp(Number(event.currentTarget.value), 0.5, 24) })}
                    aria-label="Shape stroke width"
                  />
                </label>
                <label>
                  Rotation
                  <input
                    type="number"
                    min="-180"
                    max="180"
                    value={Math.round(selected.rotation)}
                    onChange={(event) => updateMark(selected.id, { rotation: Number(event.currentTarget.value) })}
                    aria-label="Shape rotation degrees"
                  />
                </label>
              </div>
              <div className="dimension-grid">
                <label>
                  Stroke opacity
                  <input
                    type="number"
                    min="0"
                    max="1"
                    step="0.05"
                    value={selected.shapeStyle.strokeOpacity}
                    onChange={(event) => updateShapeStyle(selected.id, { strokeOpacity: clamp(Number(event.currentTarget.value), 0, 1) })}
                    aria-label="Shape stroke opacity"
                  />
                </label>
                <label>
                  Fill opacity
                  <input
                    type="number"
                    min="0"
                    max="1"
                    step="0.05"
                    value={selected.shapeStyle.fillOpacity}
                    onChange={(event) => updateShapeStyle(selected.id, { fillOpacity: clamp(Number(event.currentTarget.value), 0, 1) })}
                    aria-label="Shape fill opacity"
                  />
                </label>
              </div>
              <label>
                Dash style
                <select
                  value={selected.shapeStyle.dashStyle}
                  onChange={(event) => updateShapeStyle(selected.id, { dashStyle: event.currentTarget.value as ShapeStyle["dashStyle"] })}
                  aria-label="Shape dash style"
                >
                  {shapeDashStyles.map((style) => (
                    <option key={style} value={style}>
                      {style}
                    </option>
                  ))}
                </select>
              </label>
              {selected.shapeStyle.shapeKind === "line" || selected.shapeStyle.shapeKind === "arrow" || selected.shapeStyle.shapeKind === "doubleArrow" ? (
                <div className="dimension-grid">
                  <label>
                    Start arrow
                    <select
                      value={selected.shapeStyle.startArrowhead}
                      onChange={(event) => updateShapeStyle(selected.id, { startArrowhead: event.currentTarget.value as ShapeStyle["startArrowhead"] })}
                      aria-label="Shape start arrowhead"
                    >
                      <option value="none">none</option>
                      <option value="arrow">arrow</option>
                    </select>
                  </label>
                  <label>
                    End arrow
                    <select
                      value={selected.shapeStyle.endArrowhead}
                      onChange={(event) => updateShapeStyle(selected.id, { endArrowhead: event.currentTarget.value as ShapeStyle["endArrowhead"] })}
                      aria-label="Shape end arrowhead"
                    >
                      <option value="none">none</option>
                      <option value="arrow">arrow</option>
                    </select>
                  </label>
                </div>
              ) : null}
              <div className="dimension-grid">
                <label>
                  W
                  <input
                    type="number"
                    value={Math.round(selected.width)}
                    onChange={(event) => updateMark(selected.id, getDimensionPatch(selected, "width", Number(event.currentTarget.value)))}
                    aria-label="Shape width"
                  />
                </label>
                <label>
                  H
                  <input
                    type="number"
                    value={Math.round(selected.height)}
                    onChange={(event) => updateMark(selected.id, getDimensionPatch(selected, "height", Number(event.currentTarget.value)))}
                    aria-label="Shape height"
                  />
                </label>
              </div>
              <div className="inspector-actions">
                <button
                  className="button ghost"
                  onPointerDown={(event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    duplicateMarkById(selected.id);
                  }}
                  onClick={(event) => {
                    if (event.detail === 0) duplicateMarkById(selected.id);
                  }}
                  title="Duplicate shape"
                  aria-label="Duplicate shape"
                >
                  Duplicate
                </button>
                <button
                  className="button ghost"
                  onPointerDown={(event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    removeMarkById(selected.id);
                  }}
                  onClick={(event) => {
                    if (event.detail === 0) removeMarkById(selected.id);
                  }}
                  title="Delete shape"
                  aria-label="Delete shape"
                >
                  Delete
                </button>
              </div>
            </div>
          ) : canShowObjectInspector && selected ? (
            <div className="control-stack">
              {!isTextMarkupKind(selected.kind) ? (
                <label>
                  Content
                  <textarea
                    value={selected.text}
                    onChange={(event) => updateMark(selected.id, { text: event.target.value })}
                    aria-label="Annotation content"
                    disabled={selected.kind === "image" || selected.kind === "pngSignature"}
                  />
                </label>
              ) : null}
              <label>
                Color
                <input
                  type="color"
                  value={selected.color}
                  onChange={(event) => updateMark(selected.id, { color: event.currentTarget.value })}
                  onInput={(event) => updateMark(selected.id, { color: event.currentTarget.value })}
                  aria-label="Annotation color"
                />
              </label>
              <label>
                Colour hex
                <input
                  type="text"
                  value={selected.color}
                  onChange={(event) => {
                    const color = normalizeHexColor(event.currentTarget.value);
                    if (color) updateMark(selected.id, { color });
                  }}
                  aria-label="Annotation color hex"
                />
              </label>
              {isTextMarkupKind(selected.kind) ? (
                <>
                  <label>
                    Opacity
                    <input
                      type="range"
                      min="0.1"
                      max="1"
                      step="0.05"
                      value={selected.opacity ?? (selected.kind === "textHighlight" ? 0.38 : 1)}
                      onChange={(event) => updateMark(selected.id, { opacity: Number(event.target.value) })}
                      aria-label="Text markup opacity"
                    />
                  </label>
                  <label>
                    Opacity value
                    <input
                      type="number"
                      min="0.1"
                      max="1"
                      step="0.05"
                      value={selected.opacity ?? (selected.kind === "textHighlight" ? 0.38 : 1)}
                      onChange={(event) => updateMark(selected.id, { opacity: clamp(Number(event.currentTarget.value), 0.1, 1) })}
                      aria-label="Text markup opacity value"
                    />
                  </label>
                  {selected.kind !== "textHighlight" ? (
                    <label>
                      Thickness
                      <input
                        type="number"
                        min="0.5"
                        max="12"
                        step="0.5"
                        value={selected.thickness ?? 2}
                        onChange={(event) => updateMark(selected.id, { thickness: clamp(Number(event.currentTarget.value), 0.5, 12) })}
                        aria-label="Text markup thickness"
                      />
                    </label>
                  ) : null}
                  <div className="inspector-actions">
                    <button className="button ghost" onClick={duplicateSelectedMark} title="Duplicate text markup" aria-label="Duplicate text markup">
                      Duplicate
                    </button>
                    <button className="button ghost" onClick={removeSelectedMark} title="Delete text markup" aria-label="Delete text markup">
                      Delete
                    </button>
                  </div>
                </>
              ) : null}
              {!isTextMarkupKind(selected.kind) ? (
                <label>
                  Size
                  <input
                    type="range"
                    min="10"
                    max="42"
                    value={selected.size}
                    onChange={(event) => updateMark(selected.id, { size: Number(event.target.value) })}
                    aria-label="Annotation text size"
                  />
                </label>
              ) : null}
              {selected.kind === "image" || selected.kind === "pngSignature" ? (
                <>
                  <label>
                    Rotation
                    <input
                      type="range"
                      min="-180"
                      max="180"
                      value={selected.rotation}
                      onChange={(event) => updateMark(selected.id, { rotation: Number(event.target.value) })}
                      aria-label={selected.kind === "pngSignature" ? "Signature rotation" : "Image rotation"}
                    />
                  </label>
                  <label>
                    Degrees
                    <input
                      type="number"
                      min="-180"
                      max="180"
                      value={Math.round(selected.rotation)}
                      onChange={(event) => updateMark(selected.id, { rotation: Number(event.target.value) })}
                      aria-label={selected.kind === "pngSignature" ? "Signature rotation degrees" : "Image rotation degrees"}
                    />
                  </label>
                </>
              ) : null}
              {selected.kind === "pngSignature" ? (
                <>
                  <label>
                    Opacity
                    <input
                      type="range"
                      min="0.1"
                      max="1"
                      step="0.05"
                      value={selected.opacity ?? 1}
                      onChange={(event) => updateMark(selected.id, { opacity: Number(event.target.value) })}
                      aria-label="Signature opacity"
                    />
                  </label>
                  <label>
                    Opacity value
                    <input
                      type="number"
                      min="0.1"
                      max="1"
                      step="0.05"
                      value={selected.opacity ?? 1}
                      onChange={(event) => updateMark(selected.id, { opacity: clamp(Number(event.currentTarget.value), 0.1, 1) })}
                      aria-label="Signature opacity value"
                    />
                  </label>
                  <label className="checkbox-row">
                    <input
                      type="checkbox"
                      checked={selected.lockAspectRatio !== false}
                      onChange={(event) => updateMark(selected.id, { lockAspectRatio: event.currentTarget.checked })}
                      aria-label="Lock signature aspect ratio"
                    />
                    Lock aspect ratio
                  </label>
                  <div className="inspector-actions">
                    <button className="button ghost" onClick={duplicateSelectedMark} title="Duplicate signature" aria-label="Duplicate signature">
                      Duplicate
                    </button>
                    <button className="button ghost" onClick={removeSelectedMark} title="Delete signature" aria-label="Delete signature">
                      Delete
                    </button>
                  </div>
                </>
              ) : null}
              {selected.kind === "stroke" ? (
                <>
                  <label>
                    Opacity
                    <input
                      type="range"
                      min="0.1"
                      max="1"
                      step="0.05"
                      value={selected.strokeOpacity ?? 1}
                      onChange={(event) => updateMark(selected.id, { strokeOpacity: Number(event.target.value) })}
                      aria-label="Stroke opacity"
                    />
                  </label>
                  <label>
                    Opacity value
                    <input
                      type="number"
                      min="0.1"
                      max="1"
                      step="0.05"
                      value={selected.strokeOpacity ?? 1}
                      onChange={(event) => updateMark(selected.id, { strokeOpacity: clamp(Number(event.currentTarget.value), 0.1, 1) })}
                      aria-label="Stroke opacity value"
                    />
                  </label>
                </>
              ) : null}
              {!isTextMarkupKind(selected.kind) ? (
                <div className="dimension-grid">
                  <label>
                    W
                    <input
                      type="number"
                      value={Math.round(selected.width)}
                      onChange={(event) => updateMark(selected.id, getDimensionPatch(selected, "width", Number(event.target.value)))}
                      aria-label="Annotation width"
                    />
                  </label>
                  <label>
                    H
                    <input
                      type="number"
                      value={Math.round(selected.height)}
                      onChange={(event) => updateMark(selected.id, getDimensionPatch(selected, "height", Number(event.target.value)))}
                      aria-label="Annotation height"
                    />
                  </label>
                </div>
              ) : null}
            </div>
          ) : activeWorkspace === "import" || activeWorkspace === "assemble" || activeWorkspace === "publish" ? null : (
            <div className="empty-panel">
              {activeWorkspaceItem.icon}
              <p>{getWorkspaceEmptyInspectorMessage(activeWorkspace)}</p>
            </div>
          )}
            </div>
          </details>

          {activeWorkspace === "review" ? (
          <details className="comments-panel inspector-section" open aria-label="Comments">
            <summary>Comments</summary>
            <div className="comments-panel-body">
            <div className="comments-panel-header">
              <button
                className="link-button"
                onClick={() => setShowAllComments((value) => !value)}
                title={showAllComments ? "Show comments on current page" : "Show comments on all pages"}
                aria-label={showAllComments ? "Show comments on current page" : "Show comments on all pages"}
              >
                {showAllComments ? "All pages" : "Page only"}
              </button>
            </div>
            <div className="comment-filters" role="group" aria-label="Filter comments">
              {(["open", "resolved", "all"] as const).map((filter) => (
                <button
                  className={commentFilter === filter ? "active" : ""}
                  key={filter}
                  onClick={() => setCommentFilter(filter)}
                  title={`Show ${filter} comments`}
                  aria-label={`Show ${filter} comments`}
                  aria-pressed={commentFilter === filter}
                >
                  {filter}
                </button>
              ))}
            </div>
            <div className="comment-list">
              {visibleComments.length > 0 ? (
                visibleComments.map((mark) => (
                  <button
                    className={`comment-entry ${selectedMark === mark.id ? "active" : ""} ${mark.comment?.resolved ? "resolved" : ""}`}
                    key={mark.id}
                    onClick={() => selectComment(mark)}
                    title={`Go to comment on page ${mark.page}`}
                    aria-label={`Go to ${mark.comment?.resolved ? "resolved" : "open"} comment on page ${mark.page}: ${getCommentPreview(mark.comment?.body ?? "")}`}
                  >
                    <span className="comment-entry-dot" style={{ backgroundColor: mark.comment?.color }} />
                    <span className="comment-entry-body">
                      <strong>Page {mark.page}</strong>
                      <span>{getCommentPreview(mark.comment?.body ?? "")}</span>
                      <small>
                        {mark.comment?.resolved ? "Resolved" : "Open"} · {mark.comment?.replies.length ?? 0} replies
                      </small>
                    </span>
                  </button>
                ))
              ) : (
                <p className="muted-text">No comments match this view.</p>
              )}
            </div>
            </div>
          </details>
          ) : null}

          <div className="document-stats">
            <div>
              <strong>{pages.length}</strong>
              <span>Pages</span>
            </div>
            <div>
              <strong>{marks.length}</strong>
              <span>Edits</span>
            </div>
          </div>
        </aside>
        ) : (
          <button className="right-panel-reopen" onClick={() => setIsRightPanelCollapsed(false)} title="Show properties panel" aria-label="Show properties panel">
            <PanelRightOpen size={18} />
          </button>
        )}
      </section>
      <footer className="statusbar" aria-label="Document status">
        <div className="statusbar-group">
          <button className="status-button" onClick={() => setCurrentPage((page) => Math.max(1, page - 1))} disabled={currentPage <= 1 || isBusy} title="Previous page" aria-label="Previous page">
            <ChevronLeft size={15} />
          </button>
          <span>
            Page <strong>{currentPage}</strong> of {pages.length}
          </span>
          <button className="status-button" onClick={() => setCurrentPage((page) => Math.min(pages.length, page + 1))} disabled={currentPage >= pages.length || isBusy} title="Next page" aria-label="Next page">
            <ChevronRight size={15} />
          </button>
        </div>
        <div className="statusbar-group">
          <button className="status-button" onClick={() => setZoom((value) => Math.max(0.55, value - 0.1))} disabled={isBusy} title="Zoom out" aria-label="Zoom out">
            <ZoomOut size={15} />
          </button>
          <span className="status-zoom">{Math.round(zoom * 100)}%</span>
          <button className="status-button" onClick={() => setZoom((value) => Math.min(1.6, value + 0.1))} disabled={isBusy} title="Zoom in" aria-label="Zoom in">
            <ZoomIn size={15} />
          </button>
          <button className="status-text-button" onClick={() => setZoom(1)} disabled={isBusy} title="Fit page" aria-label="Fit page">
            Fit page
          </button>
          <button className="status-text-button" onClick={() => setZoom(1.25)} disabled={isBusy} title="Fit width" aria-label="Fit width">
            Fit width
          </button>
        </div>
        <div className="statusbar-meta">
          {saveStatus} | {activeWorkspaceItem.title} | {marks.length} edits | {comments.length} comments | {navigationIssues.length + numberingIssues.length} warnings
        </div>
      </footer>
      <ToastLayer toasts={toasts} onDismiss={(id) => setToasts((existing) => existing.filter((toast) => toast.id !== id))} />
    </main>
  );
}

function ImportSourceManager({
  sourceDocuments,
  pages,
  docxImportGroups,
  pptxImportGroups,
  isBusy,
  onReviewPage,
  onReimportDocx,
  onReimportPptx,
}: {
  sourceDocuments: Record<string, SourceDocument>;
  pages: PageView[];
  docxImportGroups: Array<{ metadata: DocxImportMetadata; pages: PageView[]; warnings: number; mappings: number }>;
  pptxImportGroups: Array<{ metadata: PptxImportMetadata; pages: PageView[]; warnings: number; mappings: number; notes: number }>;
  isBusy: boolean;
  onReviewPage: (page: PageView) => void;
  onReimportDocx: (metadata: DocxImportMetadata) => void;
  onReimportPptx: (metadata: PptxImportMetadata) => void;
}) {
  const pdfSources = Object.values(sourceDocuments).filter((source) => (source.mimeType ?? "application/pdf") === "application/pdf");
  const sourceCount = pdfSources.length + docxImportGroups.length + pptxImportGroups.length;
  return (
    <section className="source-manager" aria-label="Import source manager">
      <div className="section-heading">
        <strong>Sources</strong>
        <span>{sourceCount} connected</span>
      </div>
      {sourceCount === 0 ? (
        <p className="muted-text">No external sources imported yet.</p>
      ) : null}
      {pdfSources.map((source) => {
        const contributedPages = pages.filter((page) => page.sourceDocumentId === source.id);
        return (
          <div className="source-row" key={source.id}>
            <span className="asset-type">PDF</span>
            <strong>{source.name}</strong>
            <small>{contributedPages.length} pages · {formatAssetSize(source.bytes.byteLength)}</small>
            <div className="asset-actions">
              <button onClick={() => contributedPages[0] && onReviewPage(contributedPages[0])} disabled={!contributedPages[0]}>View pages</button>
            </div>
          </div>
        );
      })}
      {docxImportGroups.map((group) => (
        <div className="source-row" key={group.metadata.importId}>
          <span className="asset-type">DOCX</span>
          <strong>{group.metadata.sourceName}</strong>
          <small>{group.pages.length} pages · {group.metadata.fidelityMode} · {group.warnings} warnings</small>
          <div className="asset-actions">
            <button onClick={() => group.pages[0] && onReviewPage(group.pages[0])} disabled={!group.pages[0]}>View</button>
            <button onClick={() => onReimportDocx(group.metadata)} disabled={isBusy || !group.metadata.originalSourceDocumentId}>Re-import</button>
          </div>
        </div>
      ))}
      {pptxImportGroups.map((group) => (
        <div className="source-row" key={group.metadata.importId}>
          <span className="asset-type">PPTX</span>
          <strong>{group.metadata.sourceName}</strong>
          <small>{group.pages.length} slides · {group.metadata.fidelityMode} · {group.notes} notes · {group.warnings} warnings</small>
          <div className="asset-actions">
            <button onClick={() => group.pages[0] && onReviewPage(group.pages[0])} disabled={!group.pages[0]}>View</button>
            <button onClick={() => onReimportPptx(group.metadata)} disabled={isBusy || !group.metadata.originalSourceDocumentId}>Re-import</button>
          </div>
        </div>
      ))}
    </section>
  );
}

function ToastLayer({ toasts, onDismiss }: { toasts: ToastMessage[]; onDismiss: (id: string) => void }) {
  if (toasts.length === 0) return null;
  return (
    <div className="toast-layer" role="region" aria-label="Notifications">
      {toasts.map((toast) => (
        <div className={`toast toast-${toast.tone}`} key={toast.id} role={toast.tone === "error" ? "alert" : "status"} aria-live={toast.tone === "error" ? "assertive" : "polite"}>
          <span className="toast-indicator" aria-hidden="true" />
          <span className="toast-content">
            <strong>{toast.title}</strong>
            {toast.detail ? <small>{toast.detail}</small> : null}
          </span>
          <button type="button" onClick={() => onDismiss(toast.id)} aria-label="Dismiss notification" title="Dismiss notification">x</button>
        </div>
      ))}
    </div>
  );
}

function ShortcutReferenceDialog({ isMac, onClose }: { isMac: boolean; onClose: () => void }) {
  const mod = isMac ? "Cmd" : "Ctrl";
  const shortcuts = [
    [`${mod}+1`, "Import workspace"],
    [`${mod}+2`, "Assemble workspace"],
    [`${mod}+3`, "Review workspace"],
    [`${mod}+4`, "Publish workspace"],
    [`${mod}+N`, "New project"],
    [`${mod}+O`, "Open project"],
    [`${mod}+S`, "Save project"],
    [`${mod}+Shift+S`, "Save project as"],
    [`${mod}+Z`, "Undo"],
    [`${mod}+Shift+Z`, "Redo"],
    [`${isMac ? "Cmd" : "Ctrl"}+Y`, "Redo"],
    [`${mod}+F`, "Search"],
    [`${mod}+Shift+P`, "Open Publish workflow"],
    ["Delete / Backspace", "Delete selected object"],
    ["Escape", "Cancel tool or clear selection"],
    ["?", "Show shortcuts"],
  ];
  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget) onClose();
    }}>
      <section className="modal-card shortcut-dialog" role="dialog" aria-modal="true" aria-labelledby="shortcut-reference-title">
        <div className="modal-header">
          <h2 id="shortcut-reference-title">Keyboard Shortcuts</h2>
          <button className="panel-icon-button" onClick={onClose} aria-label="Close shortcut reference" title="Close dialog">x</button>
        </div>
        <div className="dialog-body">
          <p className="dialog-note">Shortcuts are disabled while typing in text fields.</p>
          <div className="shortcut-grid">
            {shortcuts.map(([keys, action]) => (
              <div className="shortcut-row" key={`${keys}-${action}`}>
                <kbd>{keys}</kbd>
                <span>{action}</span>
              </div>
            ))}
          </div>
          <div className="dialog-actions">
            <button className="button primary" onClick={onClose}>Done</button>
          </div>
        </div>
      </section>
    </div>
  );
}

function DocumentPage({
  page,
  marks,
  zoom,
  selectedMark,
  activeTool,
  penColor,
  penWidth,
  penOpacity,
  pendingSignature,
  publishingPreviewItems,
  publishingImageAssets,
  onSelect,
  onAddMark,
  onAddTextBox,
  onAddShape,
  onAddStroke,
  onAddTextMarkup,
  onPlaceSignature,
  onPreviewMark,
  onCommitMarkChange,
  editingText,
  onStartTextEdit,
  onTextDraftChange,
  onSaveTextEdit,
  onCancelTextEdit,
}: {
  page: PageView;
  marks: Mark[];
  zoom: number;
  selectedMark: string | null;
  activeTool: Tool;
  penColor: string;
  penWidth: number;
  penOpacity: number;
  pendingSignature: PreparedImage | null;
  publishingPreviewItems: ReturnType<typeof getPublishingPreviewItems>;
  publishingImageAssets: Array<{ id: string; name: string; dataUrl?: string; mimeType?: string }>;
  onSelect: (id: string | null) => void;
  onAddMark: (page: number, x: number, y: number) => void;
  onAddTextBox: (page: number, x: number, y: number, width?: number, height?: number) => void;
  onAddShape: (page: number, placement: { x: number; y: number; width?: number; height?: number }, shapeKind: ShapeKind) => void;
  onAddStroke: (page: number, points: StrokePoint[]) => void;
  onAddTextMarkup: (page: number, rects: MarkupRect[]) => void;
  onPlaceSignature: (page: number, x: number, y: number, width?: number, height?: number) => void;
  onPreviewMark: (id: string, patch: Partial<Mark>) => void;
  onCommitMarkChange: (beforeMark: Mark, afterMark: Mark) => void;
  editingText: { id: string; before: Mark; draft: string; isNew: boolean } | null;
  onStartTextEdit: (mark: Mark) => void;
  onTextDraftChange: (draft: string) => void;
  onSaveTextEdit: () => void;
  onCancelTextEdit: () => void;
}) {
  const pageRef = useRef<HTMLDivElement>(null);
  const textEditorRef = useRef<HTMLTextAreaElement>(null);
  const dragRef = useRef<{
    before: Mark;
    latest: Mark;
    startX: number;
    startY: number;
    mode: "move" | "resize" | "leader";
    corner?: "nw" | "ne" | "sw" | "se";
  } | null>(null);
  const [draftStroke, setDraftStroke] = useState<StrokePoint[]>([]);
  const [draftSignature, setDraftSignature] = useState<{ start: StrokePoint; current: StrokePoint } | null>(null);
  const [draftTextSelection, setDraftTextSelection] = useState<{ start: StrokePoint; current: StrokePoint } | null>(null);
  const [draftTextBox, setDraftTextBox] = useState<{ start: StrokePoint; current: StrokePoint } | null>(null);
  const [draftShape, setDraftShape] = useState<{ start: StrokePoint; current: StrokePoint; shapeKind: ShapeKind } | null>(null);
  const activeDraftStroke = Array.isArray(draftStroke) ? draftStroke : [];

  useEffect(() => {
    if (!editingText) return;
    window.setTimeout(() => {
      const editor = textEditorRef.current;
      editor?.focus();
      editor?.setSelectionRange(editor.value.length, editor.value.length);
    }, 0);
  }, [editingText?.id]);

  function pagePoint(event: PointerEvent<HTMLDivElement>) {
    const bounds = pageRef.current?.getBoundingClientRect();
    if (!bounds) return { x: 0, y: 0 };
    return {
      x: clamp((event.clientX - bounds.left) / zoom, 0, page.width),
      y: clamp((event.clientY - bounds.top) / zoom, 0, page.height),
    };
  }

  function createMarkupFromSelection() {
    if (!isTextMarkupKind(activeTool)) return;
    const selection = window.getSelection();
    const pageBounds = pageRef.current?.getBoundingClientRect();
    if (!selection || selection.rangeCount === 0 || selection.isCollapsed || !pageBounds) return;

    const rects = normalizeSelectionRects(selection.getRangeAt(0).getClientRects(), pageBounds, zoom, page);
    selection.removeAllRanges();
    if (rects.length === 0) return;
    onAddTextMarkup(page.pageNumber, rects);
  }

  function createMarkupFromDrag(start: StrokePoint, current: StrokePoint) {
    if (!isTextMarkupKind(activeTool)) return;
    if (Math.hypot(current.x - start.x, current.y - start.y) < 8) return;

    const minX = Math.min(start.x, current.x);
    const minY = Math.min(start.y, current.y);
    const maxX = Math.max(start.x, current.x);
    const maxY = Math.max(start.y, current.y);
    const rects = page.textItems
      .map((item) => {
        const itemRight = item.x + item.width;
        const itemBottom = item.y + item.height;
        const intersects = itemRight >= minX && item.x <= maxX && itemBottom >= minY && item.y <= maxY;
        if (!intersects) return null;

        const x = clamp(Math.max(item.x, minX), 0, page.width);
        const y = clamp(item.y, 0, page.height);
        const width = clamp(Math.min(itemRight, maxX) - x, 0, page.width - x);
        return { x, y, width, height: item.height };
      })
      .filter((rect): rect is MarkupRect => Boolean(rect && rect.width > 3 && rect.height > 3));

    if (rects.length > 0) {
      onAddTextMarkup(page.pageNumber, rects);
      return;
    }

    if (activeTool === "strikethrough") {
      onAddTextMarkup(page.pageNumber, [
        {
          x: clamp(minX, 0, page.width),
          y: clamp(minY, 0, page.height),
          width: clamp(maxX - minX, 0, page.width - minX),
          height: clamp(maxY - minY, 8, page.height - minY),
        },
      ]);
    }
  }

  return (
    <div
      ref={pageRef}
      className={`document-page ${isTextMarkupKind(activeTool) ? "text-selection-mode" : ""}`}
      style={{ width: page.width * zoom, height: page.height * zoom }}
      onPointerDown={(event) => {
        if ((event.target as HTMLElement).closest(".mark")) return;
        const point = pagePoint(event);
        if (isTextMarkupKind(activeTool)) {
          event.currentTarget.setPointerCapture(event.pointerId);
          onSelect(null);
          setDraftTextSelection({ start: point, current: point });
          return;
        }
        if (activeTool === "draw") {
          event.currentTarget.setPointerCapture(event.pointerId);
          onSelect(null);
          setDraftStroke([point]);
          return;
        }
        if (activeTool === "text") {
          event.currentTarget.setPointerCapture(event.pointerId);
          onSelect(null);
          setDraftTextBox({ start: point, current: point });
          return;
        }
        if (isShapeTool(activeTool)) {
          event.currentTarget.setPointerCapture(event.pointerId);
          onSelect(null);
          setDraftShape({ start: point, current: point, shapeKind: getShapeKindFromTool(activeTool) });
          return;
        }
        if (activeTool === "pngSignature") {
          if (!pendingSignature) {
            onAddMark(page.pageNumber, point.x, point.y);
            return;
          }
          event.currentTarget.setPointerCapture(event.pointerId);
          onSelect(null);
          setDraftSignature({ start: point, current: point });
          return;
        }
        onSelect(null);
        onAddMark(page.pageNumber, point.x, point.y);
      }}
          onPointerMove={(event) => {
            if (activeDraftStroke.length > 0) {
              setDraftStroke((existing) => [...(Array.isArray(existing) ? existing : []), pagePoint(event)]);
              return;
            }
            if (draftSignature) {
              setDraftSignature((existing) => (existing ? { ...existing, current: pagePoint(event) } : existing));
              return;
            }
            if (draftTextSelection) {
              setDraftTextSelection((existing) => (existing ? { ...existing, current: pagePoint(event) } : existing));
              return;
            }
            if (draftTextBox) {
              setDraftTextBox((existing) => (existing ? { ...existing, current: pagePoint(event) } : existing));
              return;
            }
            if (draftShape) {
              setDraftShape((existing) => (existing ? { ...existing, current: pagePoint(event) } : existing));
              return;
            }
            const drag = dragRef.current;
            if (!drag) return;
            const activeMark = marks.find((mark) => mark.id === drag.before.id);
            if (!activeMark) return;
            const dx = (event.clientX - drag.startX) / zoom;
            const dy = (event.clientY - drag.startY) / zoom;
            const latest =
              drag.mode === "leader"
                ? moveCalloutLeaderBy(drag.before, dx, dy, [page])
                : drag.mode === "resize" && drag.corner
                ? resizeMarkFromCorner(drag.before, drag.corner, dx, dy, [page])
                : moveMarkBy(drag.before, dx, dy, [page]);
            dragRef.current = { ...drag, latest };
            onPreviewMark(activeMark.id, latest);
          }}
      onPointerUp={(event) => {
        if (activeDraftStroke.length > 0) {
          event.currentTarget.releasePointerCapture(event.pointerId);
          const nextStroke = [...activeDraftStroke, pagePoint(event)];
          setDraftStroke([]);
          onAddStroke(page.pageNumber, nextStroke);
          return;
        }
        if (isTextMarkupKind(activeTool)) {
          if (draftTextSelection) {
            event.currentTarget.releasePointerCapture(event.pointerId);
            createMarkupFromDrag(draftTextSelection.start, pagePoint(event));
            setDraftTextSelection(null);
            window.getSelection()?.removeAllRanges();
            return;
          }
          window.setTimeout(createMarkupFromSelection, 0);
          return;
        }
        if (draftTextBox) {
          event.currentTarget.releasePointerCapture(event.pointerId);
          const placement = getDraftTextBoxPlacement(draftTextBox.start, pagePoint(event));
          setDraftTextBox(null);
          onAddTextBox(page.pageNumber, placement.x, placement.y, placement.width, placement.height);
          return;
        }
        if (draftShape) {
          event.currentTarget.releasePointerCapture(event.pointerId);
          const placement = getDraftTextBoxPlacement(draftShape.start, pagePoint(event));
          setDraftShape(null);
          onAddShape(page.pageNumber, placement, draftShape.shapeKind);
          return;
        }
        if (draftSignature && pendingSignature) {
          event.currentTarget.releasePointerCapture(event.pointerId);
          const placement = getSignaturePlacement(draftSignature.start, pagePoint(event), pendingSignature);
          setDraftSignature(null);
          onPlaceSignature(page.pageNumber, placement.x, placement.y, placement.width, placement.height);
          return;
        }
        const drag = dragRef.current;
        dragRef.current = null;
        if (drag && !areMarksEqual([drag.before], [drag.latest])) {
          onCommitMarkChange(drag.before, drag.latest);
        }
      }}
      onPointerLeave={() => {
        if (activeDraftStroke.length > 0) return;
        if (draftTextSelection) return;
        if (draftTextBox) return;
        if (draftShape) return;
        const drag = dragRef.current;
        dragRef.current = null;
        if (drag && !areMarksEqual([drag.before], [drag.latest])) {
          onCommitMarkChange(drag.before, drag.latest);
        }
      }}
    >
      {page.imageUrl ? <img className="pdf-image" src={page.imageUrl} alt={`Page ${page.pageNumber}`} /> : <BlankPage />}
      {page.textItems.length > 0 ? (
        <div className="text-layer" aria-hidden="true">
          {page.textItems.map((item) => (
            <span
              key={item.id}
              style={{
                left: item.x * zoom,
                top: item.y * zoom,
                width: item.width * zoom,
                height: item.height * zoom,
                fontSize: item.fontSize * zoom,
              }}
            >
              {item.text}
            </span>
          ))}
        </div>
      ) : null}
      {publishingPreviewItems.length > 0 ? (
        <div className="publishing-preview-layer" aria-hidden="true">
          {publishingPreviewItems.map((item) => {
            const imageAsset = item.assetId ? publishingImageAssets.find((asset) => asset.id === item.assetId) : undefined;
            return item.kind === "safeArea" ? (
              <div
                key={item.id}
                className="publishing-safe-area"
                style={{
                  left: item.x * zoom,
                  top: item.y * zoom,
                  width: (item.width ?? 0) * zoom,
                  height: (item.height ?? 0) * zoom,
                }}
              />
            ) : item.kind === "image" && imageAsset?.dataUrl ? (
              <img
                key={item.id}
                className="publishing-preview-image"
                src={imageAsset.dataUrl}
                alt=""
                style={{
                  left: item.x * zoom,
                  top: item.y * zoom,
                  width: (item.width ?? 0) * zoom,
                  height: (item.height ?? 0) * zoom,
                  opacity: item.opacity ?? 1,
                  transform: `translate(${item.align === "center" ? "-50%" : item.align === "right" ? "-100%" : "0"}, -50%)`,
                }}
              />
            ) : item.kind === "missingImage" ? (
              <div
                key={item.id}
                className="publishing-missing-image"
                style={{
                  left: item.x * zoom,
                  top: item.y * zoom,
                  width: (item.width ?? 0) * zoom,
                  height: (item.height ?? 0) * zoom,
                  transform: `translate(${item.align === "center" ? "-50%" : item.align === "right" ? "-100%" : "0"}, -50%)`,
                }}
              >
                Missing image
              </div>
            ) : (
              <div
                key={item.id}
                className={`publishing-preview-item publishing-${item.kind}`}
                style={{
                  left: item.x * zoom,
                  top: item.y * zoom,
                  color: item.style?.color,
                  opacity: item.opacity ?? item.style?.opacity,
                  fontSize: (item.style?.fontSize ?? 10) * zoom,
                  fontWeight: item.style?.bold ? 800 : 600,
                  fontStyle: item.style?.italic ? "italic" : undefined,
                  fontFamily: getPublishingCssFontFamily(item.style?.fontFamily),
                  textAlign: item.align,
                  transform: `translate(${item.align === "center" ? "-50%" : item.align === "right" ? "-100%" : "0"}, -50%) rotate(${item.rotation ?? 0}deg)`,
                }}
              >
                {item.text}
              </div>
            );
          })}
        </div>
      ) : null}
      {draftTextSelection ? (
        <div
          className="text-selection-preview"
          aria-hidden="true"
          style={getDraftSelectionStyle(draftTextSelection.start, draftTextSelection.current, zoom)}
        />
      ) : null}
      {activeDraftStroke.length > 0 ? (
        <svg className="draft-stroke-layer" viewBox={`0 0 ${page.width} ${page.height}`} aria-hidden="true">
          <path
            d={pointsToSvgPath(smoothStrokePoints(activeDraftStroke))}
            fill="none"
            stroke={penColor}
            strokeLinecap="round"
            strokeLinejoin="round"
            strokeOpacity={penOpacity}
            strokeWidth={penWidth}
          />
        </svg>
      ) : null}
      {draftSignature && pendingSignature ? (
        <div
          className="signature-placement-preview"
          aria-hidden="true"
          style={{
            ...scalePlacement(getSignaturePlacement(draftSignature.start, draftSignature.current, pendingSignature), zoom),
          }}
        >
          <img src={pendingSignature.dataUrl} alt="" />
        </div>
      ) : null}
      {draftTextBox ? (
        <div className="text-box-placement-preview" aria-hidden="true" style={getDraftSelectionStyle(draftTextBox.start, draftTextBox.current, zoom)} />
      ) : null}
      {draftShape ? (
        <div className="shape-placement-preview" aria-hidden="true" style={getDraftSelectionStyle(draftShape.start, draftShape.current, zoom)} />
      ) : null}
      {marks.map((mark) => (
        <div
          className={`mark mark-${mark.kind} ${selectedMark === mark.id ? "selected" : ""} ${editingText?.id === mark.id ? "editing" : ""} ${
            mark.kind === "comment" && mark.comment?.resolved ? "resolved" : ""
          } ${mark.kind === "text" && doesTextOverflow(mark) ? "overflowing" : ""}`}
          key={mark.id}
          role="button"
          tabIndex={0}
          title={getMarkLabel(mark)}
          aria-label={getMarkLabel(mark)}
          style={{
            left: mark.x * zoom,
            top: mark.y * zoom,
            width: mark.kind === "comment" ? COMMENT_MARKER_SIZE : mark.width * zoom,
            height: mark.kind === "comment" ? COMMENT_MARKER_SIZE : mark.height * zoom,
            color: mark.color,
            fontSize: mark.kind === "comment" ? 16 : mark.size * zoom,
            backgroundColor: mark.kind === "highlight" ? hexToRgba(mark.color, 0.48) : undefined,
            transform: mark.rotation ? `rotate(${mark.rotation}deg)` : undefined,
            ...getTextBoxContainerStyle(mark, zoom),
          }}
          onPointerDown={(event) => {
            if (editingText?.id === mark.id) return;
            event.stopPropagation();
            onSelect(mark.id);
            dragRef.current = {
              before: mark,
              latest: mark,
              startX: event.clientX,
              startY: event.clientY,
              mode: "move",
            };
          }}
          onDoubleClick={(event) => {
            if (mark.kind !== "text") return;
            event.stopPropagation();
            onStartTextEdit(mark);
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter" && mark.kind === "text" && editingText?.id !== mark.id) {
              event.preventDefault();
              onStartTextEdit(mark);
            }
          }}
        >
          {mark.kind === "text" ? (
            editingText?.id === mark.id ? (
              <textarea
                ref={textEditorRef}
                className="text-box-editor"
                value={editingText.draft}
                onChange={(event) => onTextDraftChange(event.currentTarget.value)}
                onPointerDown={(event) => event.stopPropagation()}
                onClick={(event) => event.stopPropagation()}
                onBlur={(event) => {
                  if (event.currentTarget.dataset.canceling === "true") return;
                  onSaveTextEdit();
                }}
                onKeyDown={(event) => {
                  event.stopPropagation();
                  if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
                    event.preventDefault();
                    onSaveTextEdit();
                  }
                  if (event.key === "Escape") {
                    event.preventDefault();
                    event.currentTarget.dataset.canceling = "true";
                    onCancelTextEdit();
                  }
                }}
                aria-label="Edit text box content"
                style={getTextBoxTextStyle(mark, zoom)}
              />
            ) : (
              <div className="text-box-preview" style={getTextBoxTextStyle(mark, zoom)}>
                {mark.text || ""}
              </div>
            )
          ) : null}
          {mark.kind === "shape" && mark.shapeStyle ? <ShapePreview mark={mark} zoom={zoom} /> : null}
          {isTextMarkupKind(mark.kind) && mark.markupRects ? (
            <svg className="markup-mark-svg" viewBox={`0 0 ${mark.width} ${mark.height}`} aria-hidden="true">
              {mark.markupRects.map((rect, index) => {
                const relativeX = rect.x - mark.x;
                const relativeY = rect.y - mark.y;
                if (mark.kind === "textHighlight") {
                  return (
                    <rect
                      key={`${mark.id}-${index}`}
                      x={relativeX}
                      y={relativeY}
                      width={rect.width}
                      height={rect.height}
                      fill={mark.color}
                      opacity={mark.opacity ?? 0.38}
                    />
                  );
                }

                const lineY = relativeY + rect.height * (mark.kind === "underline" ? 0.88 : 0.5);
                return (
                  <line
                    key={`${mark.id}-${index}`}
                    x1={relativeX}
                    y1={lineY}
                    x2={relativeX + rect.width}
                    y2={lineY}
                    stroke={mark.color}
                    strokeLinecap="round"
                    strokeWidth={mark.thickness ?? 2}
                    opacity={mark.opacity ?? 1}
                  />
                );
              })}
            </svg>
          ) : null}
          {(mark.kind === "image" || mark.kind === "pngSignature") && mark.imageDataUrl ? (
            <img
              className="image-mark-media"
              src={mark.imageDataUrl}
              alt={mark.kind === "pngSignature" ? mark.imageName || "PNG signature annotation" : mark.imageName || "Image annotation"}
              style={{ opacity: mark.opacity ?? 1 }}
            />
          ) : null}
          {mark.kind === "stroke" && mark.strokePoints ? (
            <svg className="stroke-mark-svg" viewBox={`0 0 ${mark.width} ${mark.height}`} aria-hidden="true">
              <path
                d={pointsToSvgPath(pointsRelativeToMark(smoothStrokePoints(mark.strokePoints), mark))}
                fill="none"
                stroke={mark.color}
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeOpacity={mark.strokeOpacity ?? 1}
                strokeWidth={mark.size}
              />
            </svg>
          ) : null}
          {mark.kind === "comment" ? <MessageSquare size={16} aria-hidden="true" /> : null}
          {mark.kind !== "highlight" &&
          mark.kind !== "image" &&
          mark.kind !== "pngSignature" &&
          mark.kind !== "stroke" &&
          mark.kind !== "comment" &&
          mark.kind !== "text" &&
          !isTextMarkupKind(mark.kind)
            ? mark.text
            : ""}
          {mark.kind === "shape" && mark.shapeStyle?.shapeKind === "callout" && selectedMark === mark.id ? (
            <span
              className="leader-handle"
              aria-hidden="true"
              style={{
                left: (mark.shapeStyle.leaderEnd?.x ?? mark.width / 2) * zoom - 5,
                top: (mark.shapeStyle.leaderEnd?.y ?? mark.height + 48) * zoom - 5,
              }}
              onPointerDown={(event) => {
                event.stopPropagation();
                onSelect(mark.id);
                dragRef.current = {
                  before: mark,
                  latest: mark,
                  startX: event.clientX,
                  startY: event.clientY,
                  mode: "leader",
                };
              }}
            />
          ) : null}
          {((mark.kind === "image" || mark.kind === "pngSignature" || mark.kind === "text" || mark.kind === "shape") && selectedMark === mark.id && editingText?.id !== mark.id)
            ? (["nw", "ne", "sw", "se"] as const).map((corner) => (
                <span
                  className={`resize-handle handle-${corner}`}
                  key={corner}
                  aria-hidden="true"
                  onPointerDown={(event) => {
                    event.stopPropagation();
                    onSelect(mark.id);
                    dragRef.current = {
                      before: mark,
                      latest: mark,
                      startX: event.clientX,
                      startY: event.clientY,
                      mode: "resize",
                      corner,
                    };
                  }}
                />
              ))
            : null}
        </div>
      ))}
    </div>
  );
}

function BlankPage() {
  return (
    <div className="blank-page">
      <FilePlus2 size={34} aria-hidden="true" />
      <span>New PDF page</span>
    </div>
  );
}

function ShapePreview({ mark, zoom }: { mark: Mark; zoom: number }) {
  const style = mark.shapeStyle;
  if (!style) return null;
  const strokeWidth = Math.max(1, style.strokeWidth * zoom);
  const dashArray = getDashArray(style.dashStyle, strokeWidth);
  const leaderEnd = style.leaderEnd ?? { x: mark.width / 2, y: mark.height + 48 };
  const common = {
    stroke: style.strokeColor,
    strokeWidth,
    strokeOpacity: style.strokeOpacity,
    fill: hexToRgba(style.fillColor, style.fillOpacity),
    strokeDasharray: dashArray,
  };

  if (style.shapeKind === "line" || style.shapeKind === "arrow" || style.shapeKind === "doubleArrow") {
    return (
      <svg className="shape-svg" viewBox={`0 0 ${mark.width} ${mark.height}`} aria-hidden="true">
        <defs>{getArrowMarkerDefinitions(mark.id, style.strokeColor)}</defs>
        <line
          x1={0}
          y1={mark.height / 2}
          x2={mark.width}
          y2={mark.height / 2}
          stroke={style.strokeColor}
          strokeWidth={style.strokeWidth}
          strokeOpacity={style.strokeOpacity}
          strokeDasharray={getDashArray(style.dashStyle, style.strokeWidth)}
          markerStart={style.startArrowhead === "arrow" ? `url(#arrow-start-${mark.id})` : undefined}
          markerEnd={style.endArrowhead === "arrow" ? `url(#arrow-end-${mark.id})` : undefined}
        />
      </svg>
    );
  }

  if (style.shapeKind === "ellipse") {
    return (
      <svg className="shape-svg" viewBox={`0 0 ${mark.width} ${mark.height}`} aria-hidden="true">
        <ellipse cx={mark.width / 2} cy={mark.height / 2} rx={Math.max(1, mark.width / 2 - style.strokeWidth)} ry={Math.max(1, mark.height / 2 - style.strokeWidth)} {...common} strokeWidth={style.strokeWidth} />
      </svg>
    );
  }

  if (style.shapeKind === "polygon") {
    return (
      <svg className="shape-svg" viewBox={`0 0 ${mark.width} ${mark.height}`} aria-hidden="true">
        <polygon points={`${mark.width / 2},2 ${mark.width - 2},${mark.height * 0.38} ${mark.width * 0.82},${mark.height - 2} ${mark.width * 0.18},${mark.height - 2} 2,${mark.height * 0.38}`} {...common} strokeWidth={style.strokeWidth} />
      </svg>
    );
  }

  if (style.shapeKind === "cloud") {
    return (
      <svg className="shape-svg" viewBox={`0 0 ${mark.width} ${mark.height}`} aria-hidden="true">
        <path d={getCloudPath(mark.width, mark.height)} {...common} strokeWidth={style.strokeWidth} />
      </svg>
    );
  }

  if (style.shapeKind === "callout") {
    return (
      <>
        <svg className="shape-svg callout-svg" viewBox={`0 0 ${Math.max(mark.width, leaderEnd.x)} ${Math.max(mark.height, leaderEnd.y)}`} aria-hidden="true">
          <line
            x1={mark.width * 0.5}
            y1={mark.height}
            x2={leaderEnd.x}
            y2={leaderEnd.y}
            stroke={style.strokeColor}
            strokeWidth={style.strokeWidth}
            strokeOpacity={style.strokeOpacity}
            strokeDasharray={getDashArray(style.dashStyle, style.strokeWidth)}
          />
          <rect x={style.strokeWidth / 2} y={style.strokeWidth / 2} width={Math.max(1, mark.width - style.strokeWidth)} height={Math.max(1, mark.height - style.strokeWidth)} rx={8} {...common} strokeWidth={style.strokeWidth} />
        </svg>
        <div className="callout-text">{style.calloutText}</div>
      </>
    );
  }

  return (
    <svg className="shape-svg" viewBox={`0 0 ${mark.width} ${mark.height}`} aria-hidden="true">
      <rect
        x={style.strokeWidth / 2}
        y={style.strokeWidth / 2}
        width={Math.max(1, mark.width - style.strokeWidth)}
        height={Math.max(1, mark.height - style.strokeWidth)}
        rx={style.shapeKind === "roundedRectangle" ? style.cornerRadius : 0}
        {...common}
        strokeWidth={style.strokeWidth}
      />
    </svg>
  );
}

function getArrowMarkerDefinitions(id: string, color: string) {
  return (
    <>
      <marker id={`arrow-end-${id}`} markerWidth="12" markerHeight="12" refX="10" refY="6" orient="auto" markerUnits="strokeWidth">
        <path d="M 0 0 L 12 6 L 0 12 z" fill={color} />
      </marker>
      <marker id={`arrow-start-${id}`} markerWidth="12" markerHeight="12" refX="2" refY="6" orient="auto" markerUnits="strokeWidth">
        <path d="M 12 0 L 0 6 L 12 12 z" fill={color} />
      </marker>
    </>
  );
}

function ToolButton({
  icon,
  label,
  active,
  disabled,
  onClick,
}: {
  icon: ReactElement;
  label: string;
  active?: boolean;
  disabled?: boolean;
  onClick?: () => void;
}) {
  return (
    <button className={`tool-button ${active ? "active" : ""}`} onClick={onClick} disabled={disabled} title={label} aria-label={label} aria-pressed={active}>
      {icon}
    </button>
  );
}

function ToolbarMenu({ icon, label, children }: { icon: ReactElement; label: string; children: ReactElement | ReactElement[] }) {
  return (
    <details className="toolbar-menu">
      <summary title={label} aria-label={label}>
        {icon}
        <span>{label}</span>
      </summary>
      <div className="toolbar-menu-panel" role="menu" aria-label={label}>
        {children}
      </div>
    </details>
  );
}

function MenuToolButton({
  icon,
  label,
  active,
  disabled,
  onClick,
}: {
  icon: ReactElement;
  label: string;
  active?: boolean;
  disabled?: boolean;
  onClick?: () => void;
}) {
  return (
    <button className={`menu-tool-button ${active ? "active" : ""}`} onClick={onClick} disabled={disabled} role="menuitem" title={label} aria-label={label} aria-pressed={active}>
      {icon}
      <span>{label}</span>
    </button>
  );
}

function WorkspaceButton({
  icon,
  title,
  description,
  active,
  onClick,
}: {
  icon: ReactElement;
  title: string;
  description: string;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <button className={`workspace-button ${active ? "active" : ""}`} onClick={onClick} title={`${title}: ${description}`} aria-label={`${title} workspace: ${description}`} aria-pressed={active}>
      {icon}
      <span>
        <strong>{title}</strong>
        <small>{description}</small>
      </span>
    </button>
  );
}

function BookmarkTreePanel({
  bookmarks,
  pages,
  selectedBookmarkId,
  issues,
  onSelect,
  onToggle,
  onRename,
  onMove,
  onNest,
  onOutdent,
  onDuplicate,
  onDelete,
  onRepair,
}: {
  bookmarks: Array<DocumentBookmark & { level: number }>;
  pages: PageView[];
  selectedBookmarkId: string | null;
  issues: ReturnType<typeof validateNavigation>;
  onSelect: (bookmark: DocumentBookmark) => void;
  onToggle: (bookmark: DocumentBookmark) => void;
  onRename: (bookmark: DocumentBookmark, title: string) => void;
  onMove: (direction: -1 | 1) => void;
  onNest: () => void;
  onOutdent: () => void;
  onDuplicate: (bookmarkId: string) => void;
  onDelete: (bookmarkId: string) => void;
  onRepair: (bookmarkId: string) => void;
}) {
  const issueByBookmark = new Map(issues.filter((issue) => issue.bookmarkId).map((issue) => [issue.bookmarkId, issue]));
  return (
    <div className="panel-list bookmark-tree-panel">
      {bookmarks.length > 0 ? (
        <div role="tree" aria-label="Document bookmarks">
          {bookmarks.map((bookmark) => {
            const issue = issueByBookmark.get(bookmark.id);
            return (
              <div className={`bookmark-tree-row ${selectedBookmarkId === bookmark.id ? "active" : ""} ${issue ? "broken" : ""}`} key={bookmark.id} role="treeitem" aria-level={bookmark.level} aria-selected={selectedBookmarkId === bookmark.id} style={{ paddingLeft: 8 + (bookmark.level - 1) * 16 }}>
                <button className="panel-icon-button" onClick={() => onToggle(bookmark)} aria-label={bookmark.expanded ? `Collapse ${bookmark.title}` : `Expand ${bookmark.title}`} title={bookmark.expanded ? "Collapse" : "Expand"}>
                  {bookmark.expanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                </button>
                <button className="bookmark-title-button" onClick={() => onSelect(bookmark)} title={`Go to ${formatBookmarkDestination(bookmark, pages)}`}>
                  {issue ? <AlertTriangle size={14} aria-hidden="true" /> : <BookOpen size={14} aria-hidden="true" />}
                  <span>{bookmark.title}</span>
                  <small>{formatBookmarkDestination(bookmark, pages)}</small>
                </button>
                <div className="bookmark-row-actions" aria-label={`Actions for ${bookmark.title}`}>
                  <button onClick={() => onMove(-1)} title="Move up" aria-label="Move bookmark up"><ChevronUp size={13} /></button>
                  <button onClick={() => onMove(1)} title="Move down" aria-label="Move bookmark down"><ChevronDown size={13} /></button>
                  <button onClick={onOutdent} title="Outdent" aria-label="Outdent bookmark">Out</button>
                  <button onClick={onNest} title="Nest under previous bookmark" aria-label="Nest bookmark">Nest</button>
                  <button onClick={() => {
                    const title = window.prompt("Bookmark title", bookmark.title)?.trim();
                    if (title) onRename(bookmark, title);
                  }} title="Rename bookmark" aria-label="Rename bookmark">Rename</button>
                  <button onClick={() => onDuplicate(bookmark.id)} title="Duplicate bookmark" aria-label="Duplicate bookmark">Copy</button>
                  {issue ? <button onClick={() => onRepair(bookmark.id)} title="Repair destination to current page" aria-label="Repair bookmark destination">Repair</button> : null}
                  <button onClick={() => onDelete(bookmark.id)} title="Delete bookmark" aria-label="Delete bookmark">Delete</button>
                </div>
                {issue ? <span className="field-error" role="status">{issue.message}</span> : null}
              </div>
            );
          })}
        </div>
      ) : (
        <div className="panel-empty">
          <BookOpen size={22} />
          <strong>No bookmarks yet</strong>
          <span>Create bookmarks from the Assemble workspace.</span>
        </div>
      )}
    </div>
  );
}

type WelcomeScreenProps = {
  logoSrc: string;
  recentProjects: RecentProject[];
  autosave: ProjectAutosave | null;
  onNewProject: () => void;
  onOpenProject: () => void;
  onOpenPdf: () => void;
  onOpenDocx: () => void;
  onOpenPptx: () => void;
  onOpenRecent: (project: RecentProject) => void;
  onRemoveRecent: (id: string) => void;
  onClearRecent: () => void;
  onRestoreAutosave: () => void;
};

function DocxImportDialog({
  file,
  mode,
  options,
  reimportTarget,
  reimportImpact,
  hasOpenProject,
  isBusy,
  onModeChange,
  onOptionsChange,
  onReplacementFile,
  onCancel,
  onImport,
}: {
  file: File;
  mode: "replace" | "append";
  options: DocxImportOptions;
  reimportTarget: DocxImportMetadata | null;
  reimportImpact: DocxReimportImpact | null;
  hasOpenProject: boolean;
  isBusy: boolean;
  onModeChange: (mode: "replace" | "append") => void;
  onOptionsChange: (options: DocxImportOptions) => void;
  onReplacementFile: (file: File) => void;
  onCancel: () => void;
  onImport: () => void;
}) {
  const update = (patch: Partial<DocxImportOptions>) => onOptionsChange({ ...options, ...patch });
  return (
    <div className="modal-backdrop" role="presentation">
      <section className="modal-card" role="dialog" aria-modal="true" aria-labelledby="docx-import-title">
        <div className="modal-header">
          <h2 id="docx-import-title">Import Word document</h2>
          <button className="panel-icon-button" onClick={onCancel} aria-label="Close Word import dialog" title="Close dialog">x</button>
        </div>
        <div className="dialog-body">
          <p className="dialog-note">
            {reimportTarget
              ? "Publish Pro will re-import this preserved Word source and replace the previous generated Word pages only after conversion succeeds."
              : "Publish Pro converts this DOCX into editable project pages. The original Word file can be preserved inside the project for traceability."}
          </p>
          <div className="import-summary">
            <strong>{file.name}</strong>
            <span>{formatAssetSize(file.size)} · DOCX</span>
          </div>
          {reimportTarget ? (
            <div className="validation-list">
              <strong>Re-import impact preview</strong>
              <div className="validation-row"><span>{reimportImpact?.currentPages ?? 0} imported pages will be replaced after conversion succeeds.</span></div>
              <div className="validation-row"><span>{reimportImpact?.currentBookmarks ?? 0} imported bookmarks will be regenerated; {reimportImpact?.manualBookmarksPreserved ?? 0} manual bookmarks are preserved.</span></div>
              <div className="validation-row"><span>{reimportImpact?.manualAnnotationsAffected ?? 0} manual annotations and {reimportImpact?.commentsAffected ?? 0} comments are remapped by page order where possible.</span></div>
              <div className="validation-row"><span>Publishing settings, project metadata, and unrelated assets are preserved.</span></div>
              <label>
                Replacement DOCX
                <input
                  type="file"
                  accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  onChange={(event) => {
                    const replacement = event.currentTarget.files?.[0];
                    event.currentTarget.value = "";
                    if (replacement) onReplacementFile(replacement);
                  }}
                  disabled={isBusy}
                />
              </label>
            </div>
          ) : null}
          <label>
            Fidelity
            <select value={options.fidelityMode} onChange={(event) => update({ fidelityMode: event.currentTarget.value as DocxImportOptions["fidelityMode"] })}>
              <option value="fast">Fast - simplified layout</option>
              <option value="balanced">Balanced - recommended</option>
              <option value="high">High Fidelity - slower, more detailed</option>
            </select>
          </label>
          {options.fidelityMode === "high" ? <p className="dialog-note">High Fidelity improves table, spacing, and image handling, but Word-specific layout may still differ.</p> : null}
          {hasOpenProject ? (
            <label>
              Import mode
              <select value={mode} onChange={(event) => onModeChange(event.currentTarget.value as "replace" | "append")} disabled={Boolean(reimportTarget)}>
                <option value="append">Append to current project</option>
                <option value="replace">Create new project from Word document</option>
              </select>
            </label>
          ) : null}
          <div className="checkbox-stack">
            <label className="checkbox-row"><input type="checkbox" checked={options.preserveSource} onChange={(event) => update({ preserveSource: event.currentTarget.checked })} />Preserve original DOCX in project sources</label>
            <label className="checkbox-row"><input type="checkbox" checked={options.createBookmarks} onChange={(event) => update({ createBookmarks: event.currentTarget.checked })} />Create bookmarks from Word headings</label>
            <label className="checkbox-row"><input type="checkbox" checked={options.importHeadersFooters} onChange={(event) => update({ importHeadersFooters: event.currentTarget.checked })} />Import simple headers and footers as publishing settings</label>
            <label className="checkbox-row"><input type="checkbox" checked={options.importHyperlinks} onChange={(event) => update({ importHyperlinks: event.currentTarget.checked })} />Preserve hyperlink text styling</label>
          </div>
          <label>
            Tracked changes
            <select value={options.trackedChangesMode} onChange={(event) => update({ trackedChangesMode: event.currentTarget.value as DocxImportOptions["trackedChangesMode"] })}>
              <option value="acceptAll">Accept all</option>
              <option value="rejectAll">Reject all</option>
              <option value="summary">Preserve review summary</option>
            </select>
          </label>
          <div className="page-field-grid">
            <label>
              Fallback page size
              <select value={options.fallbackPageSize} onChange={(event) => update({ fallbackPageSize: event.currentTarget.value as DocxImportOptions["fallbackPageSize"] })}>
                <option value="a4">A4</option>
                <option value="letter">Letter</option>
                <option value="legal">Legal</option>
              </select>
            </label>
            <label>
              Fallback font
              <select value={options.fallbackFont} onChange={(event) => update({ fallbackFont: event.currentTarget.value as DocxImportOptions["fallbackFont"] })}>
                <option value="Helvetica">Helvetica</option>
                <option value="Times Roman">Times Roman</option>
                <option value="Courier">Courier</option>
              </select>
            </label>
          </div>
          <p className="dialog-note">This first phase imports Office Open XML content locally in the browser. Legacy `.doc` files are not supported.</p>
          <div className="dialog-actions">
            <button className="button ghost" onClick={onCancel} disabled={isBusy}>Cancel</button>
            <button className="button primary" onClick={onImport} disabled={isBusy}>{reimportTarget ? "Re-import Word document" : "Import Word document"}</button>
          </div>
        </div>
      </section>
    </div>
  );
}

function normalizeDocxImportOptions(options: DocxImportOptions): DocxImportOptions {
  const legacyMode = options.trackedChangesMode as DocxImportOptions["trackedChangesMode"] | "accepted" | "rejectDeletions";
  return {
    ...defaultDocxImportOptions,
    ...options,
    trackedChangesMode: legacyMode === "accepted" ? "acceptAll" : legacyMode === "rejectDeletions" ? "rejectAll" : legacyMode,
  };
}

function isPptxImportMetadata(metadata: OfficeImportMetadata): metadata is PptxImportMetadata {
  return (metadata as PptxImportMetadata).kind === "pptx";
}

function PptxImportDialog({
  file,
  mode,
  options,
  reimportTarget,
  reimportImpact,
  hasOpenProject,
  isBusy,
  onModeChange,
  onOptionsChange,
  onReplacementFile,
  onCancel,
  onImport,
}: {
  file: File;
  mode: "replace" | "append";
  options: PptxImportOptions;
  reimportTarget: PptxImportMetadata | null;
  reimportImpact: PptxReimportImpact | null;
  hasOpenProject: boolean;
  isBusy: boolean;
  onModeChange: (mode: "replace" | "append") => void;
  onOptionsChange: (options: PptxImportOptions) => void;
  onReplacementFile: (file: File) => void;
  onCancel: () => void;
  onImport: () => void;
}) {
  const update = (patch: Partial<PptxImportOptions>) => onOptionsChange({ ...options, ...patch });
  return (
    <div className="modal-backdrop" role="presentation">
      <section className="modal-card" role="dialog" aria-modal="true" aria-labelledby="pptx-import-title">
        <div className="modal-header">
          <h2 id="pptx-import-title">Import PowerPoint presentation</h2>
          <button className="panel-icon-button" onClick={onCancel} aria-label="Close PowerPoint import dialog" title="Close dialog">x</button>
        </div>
        <div className="dialog-body">
          <p className="dialog-note">
            {reimportTarget
              ? "Publish Pro will re-import this preserved PowerPoint source and replace only the previous generated slide pages after conversion succeeds."
              : "Publish Pro converts PPTX slides into project pages. This is slide-to-page publishing import, not full PowerPoint editing."}
          </p>
          <div className="import-summary">
            <strong>{file.name}</strong>
            <span>{formatAssetSize(file.size)} · PPTX</span>
          </div>
          {reimportTarget ? (
            <div className="validation-list">
              <strong>Re-import impact preview</strong>
              <div className="validation-row"><span>{reimportImpact?.currentSlides ?? 0} imported slides will be replaced after conversion succeeds.</span></div>
              <div className="validation-row"><span>{reimportImpact?.currentBookmarks ?? 0} slide-title bookmarks will be regenerated.</span></div>
              <div className="validation-row"><span>{reimportImpact?.manualAnnotationsAffected ?? 0} manual annotations are remapped by slide order where possible.</span></div>
              <div className="validation-row"><span>{reimportImpact?.speakerNotes ?? 0} speaker notes are tracked in metadata. Publishing settings are preserved.</span></div>
              <label>
                Replacement PPTX
                <input
                  type="file"
                  accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
                  onChange={(event) => {
                    const replacement = event.currentTarget.files?.[0];
                    event.currentTarget.value = "";
                    if (replacement) onReplacementFile(replacement);
                  }}
                  disabled={isBusy}
                />
              </label>
            </div>
          ) : null}
          <label>
            Fidelity
            <select value={options.fidelityMode} onChange={(event) => update({ fidelityMode: event.currentTarget.value as PptxImportOptions["fidelityMode"] })}>
              <option value="fast">Fast - simplified slide rendering</option>
              <option value="balanced">Balanced - recommended</option>
              <option value="high">High Fidelity - preserves more metadata</option>
            </select>
          </label>
          {hasOpenProject ? (
            <label>
              Import mode
              <select value={mode} onChange={(event) => onModeChange(event.currentTarget.value as "replace" | "append")} disabled={Boolean(reimportTarget)}>
                <option value="append">Import into current project</option>
                <option value="replace">Create project from PowerPoint</option>
              </select>
            </label>
          ) : null}
          <div className="checkbox-stack">
            <label className="checkbox-row"><input type="checkbox" checked={options.preserveSource} onChange={(event) => update({ preserveSource: event.currentTarget.checked })} />Preserve original PPTX in project sources</label>
            <label className="checkbox-row"><input type="checkbox" checked={options.createBookmarks} onChange={(event) => update({ createBookmarks: event.currentTarget.checked })} />Create bookmarks from slide titles</label>
            <label className="checkbox-row"><input type="checkbox" checked={options.includeHiddenSlides} onChange={(event) => update({ includeHiddenSlides: event.currentTarget.checked })} />Include hidden slides and mark them clearly</label>
          </div>
          <p className="dialog-note">Legacy `.ppt` files, animations, transitions, and full PowerPoint editing are not supported in this phase.</p>
          <div className="dialog-actions">
            <button className="button ghost" onClick={onCancel} disabled={isBusy}>Cancel</button>
            <button className="button primary" onClick={onImport} disabled={isBusy}>{reimportTarget ? "Re-import PowerPoint" : "Import PowerPoint"}</button>
          </div>
        </div>
      </section>
    </div>
  );
}

function PptxImportReportDialog({ report, onClose }: { report: PptxImportReport; onClose: () => void }) {
  const categories = Array.from(new Set(report.warnings.map((warning) => warning.category ?? "general")));
  const [categoryFilter, setCategoryFilter] = useState<string>("all");
  const visibleWarnings = categoryFilter === "all" ? report.warnings : report.warnings.filter((warning) => (warning.category ?? "general") === categoryFilter);
  const copyReport = () => {
    void navigator.clipboard?.writeText(JSON.stringify(report, null, 2));
  };
  const downloadReport = () => {
    const url = URL.createObjectURL(new Blob([JSON.stringify(report, null, 2)], { type: "application/json" }));
    const link = document.createElement("a");
    link.href = url;
    link.download = `publish-pro-pptx-import-${report.importId}.json`;
    link.click();
    URL.revokeObjectURL(url);
  };
  return (
    <div className="modal-backdrop" role="presentation">
      <section className="modal-card" role="dialog" aria-modal="true" aria-labelledby="pptx-report-title">
        <div className="modal-header">
          <h2 id="pptx-report-title">PowerPoint import report</h2>
          <button className="panel-icon-button" onClick={onClose} aria-label="Close PowerPoint import report" title="Close report">x</button>
        </div>
        <div className="dialog-body">
          <div className="project-stats-grid">
            <div><strong>{report.slidesImported}</strong><span>Slides</span></div>
            <div><strong>{report.hiddenSlides}</strong><span>Hidden</span></div>
            <div><strong>{report.slideSections}</strong><span>Sections</span></div>
            <div><strong>{report.textBoxes}</strong><span>Text boxes</span></div>
            <div><strong>{report.images}</strong><span>Images</span></div>
            <div><strong>{report.shapes}</strong><span>Shapes</span></div>
            <div><strong>{report.tables}</strong><span>Tables</span></div>
            <div><strong>{report.charts}</strong><span>Charts</span></div>
            <div><strong>{report.speakerNotes}</strong><span>Notes</span></div>
            <div><strong>{report.hyperlinks}</strong><span>Links</span></div>
          </div>
          <div className="import-review-summary">
            <span>{report.fidelityMode}</span>
            <span>Parse {Math.round(report.parseTimeMs)}ms</span>
            <span>Render {Math.round(report.renderTimeMs)}ms</span>
            <span>Total {Math.round(report.totalTimeMs)}ms</span>
            <span>Rev {report.revision.sourceHash.slice(0, 8)}</span>
          </div>
          {report.fontSubstitutions.length > 0 ? (
            <div className="validation-list">
              <strong>Font substitutions</strong>
              {report.fontSubstitutions.map((item, index) => <div className="validation-row warning" key={`pptx-font-${index}`}><AlertTriangle size={15} /><span>{item}</span></div>)}
            </div>
          ) : null}
          {report.warnings.length > 0 ? (
            <div className="validation-list">
              <div className="dialog-actions compact">
                <strong>Import warnings</strong>
                <select value={categoryFilter} onChange={(event) => setCategoryFilter(event.currentTarget.value)} aria-label="Filter PowerPoint import warnings">
                  <option value="all">All warnings</option>
                  {categories.map((category) => <option value={category} key={category}>{category}</option>)}
                </select>
              </div>
              {visibleWarnings.map((warning, index) => (
                <div className="validation-row warning" key={`${warning.code}-${index}`}>
                  <AlertTriangle size={15} />
                  <span>{warning.message}{warning.pageNumber ? ` Slide ${warning.pageNumber}.` : ""}</span>
                </div>
              ))}
            </div>
          ) : <p className="dialog-note">Import completed without warnings.</p>}
          <div className="dialog-actions">
            <button className="button ghost" onClick={copyReport}>Copy report</button>
            <button className="button ghost" onClick={downloadReport}>Download JSON</button>
            <button className="button primary" onClick={onClose}>Done</button>
          </div>
        </div>
      </section>
    </div>
  );
}

function createWordCommentMarks(wordComments: DocxWordComment[], importedPages: PageView[]): Mark[] {
  const pagesByNumber = new Map(importedPages.map((page) => [page.pageNumber, page]));
  const mappingsByBlockId = new Map(importedPages.flatMap((page) => (page.sourceMappings ?? []).flatMap((mapping) => ("blockId" in mapping ? [[mapping.blockId, { page, mapping }]] : []))));
  return wordComments.flatMap((wordComment) => {
    const mapped = wordComment.blockId ? mappingsByBlockId.get(wordComment.blockId) : undefined;
    const page = mapped?.page ?? pagesByNumber.get(1) ?? importedPages[0];
    if (!page) return [];
    const comment = createCommentData(wordComment.date ? new Date(wordComment.date) : new Date(), wordComment.author || "Word reviewer");
    comment.title = `Word comment ${wordComment.sourceId}`;
    comment.body = [
      wordComment.text,
      "",
      `Source comment ID: ${wordComment.sourceId}`,
      wordComment.initials ? `Initials: ${wordComment.initials}` : "",
      wordComment.sourceText ? `Referenced text: ${wordComment.sourceText}` : "",
      wordComment.approximate ? "Mapping: approximate" : "Mapping: source range",
    ].filter(Boolean).join("\n");
    comment.color = "#60a5fa";
    return [{
      id: crypto.randomUUID(),
      kind: "comment" as const,
      pageId: page.id,
      page: page.pageNumber,
      x: Math.max(16, Math.min(page.width - 32, mapped?.mapping.x ?? page.width - 56)),
      y: Math.max(16, Math.min(page.height - 32, mapped?.mapping.y ?? 72)),
      width: 24,
      height: 24,
      text: comment.title,
      color: comment.color,
      size: 16,
      rotation: 0,
      comment,
    }];
  });
}

function createDocxNumberingSections(sections: DocxSectionNumbering[], importedPages: PageView[]): NumberingSection[] {
  if (sections.length === 0 || importedPages.length === 0) return [];
  const firstPage = importedPages[0];
  return sections.map((section, index) => ({
    id: crypto.randomUUID(),
    label: index === 0 ? "Word numbering" : `Word numbering ${index + 1}`,
    startPageId: importedPages[Math.min(index, importedPages.length - 1)]?.id ?? firstPage.id,
    format: section.format,
    startValue: Math.max(1, section.startValue),
    prefix: section.prefix,
    suffix: "",
    includeNumbering: !section.unnumbered,
    includeInTotal: !section.unnumbered,
    restart: true,
    usePageLabel: false,
  }));
}

function mergeDocxRevisionHistory(next: DocxImportMetadata, previous: DocxImportMetadata) {
  const nextRevision = next.revisionHistory[0];
  return [nextRevision, ...(previous.revisionHistory ?? [])].filter(Boolean).slice(0, 2);
}

function DocxImportReportDialog({ report, onClose }: { report: DocxImportReport; onClose: () => void }) {
  const categories = Array.from(new Set(report.warnings.map((warning) => warning.category ?? "general")));
  const [categoryFilter, setCategoryFilter] = useState<string>("all");
  const visibleWarnings = categoryFilter === "all" ? report.warnings : report.warnings.filter((warning) => (warning.category ?? "general") === categoryFilter);
  const copyReport = () => {
    void navigator.clipboard?.writeText(JSON.stringify(report, null, 2));
  };
  const downloadReport = () => {
    const url = URL.createObjectURL(new Blob([JSON.stringify(report, null, 2)], { type: "application/json" }));
    const link = document.createElement("a");
    link.href = url;
    link.download = `publish-pro-docx-import-${report.importId}.json`;
    link.click();
    URL.revokeObjectURL(url);
  };
  return (
    <div className="modal-backdrop" role="presentation">
      <section className="modal-card" role="dialog" aria-modal="true" aria-labelledby="docx-report-title">
        <div className="modal-header">
          <h2 id="docx-report-title">Word import report</h2>
          <button className="panel-icon-button" onClick={onClose} aria-label="Close Word import report" title="Close report">x</button>
        </div>
        <div className="dialog-body">
          <div className="project-stats-grid">
            <div><strong>{report.fidelityMode}</strong><span>Fidelity</span></div>
            <div><strong>{report.pagesCreated}</strong><span>Pages</span></div>
            <div><strong>{report.sectionsDetected}</strong><span>Sections</span></div>
            <div><strong>{report.headingsFound}</strong><span>Headings</span></div>
            <div><strong>{report.bookmarksCreated}</strong><span>Bookmarks</span></div>
            <div><strong>{report.imagesImported}</strong><span>Images</span></div>
            <div><strong>{report.tablesImported}</strong><span>Tables</span></div>
            <div><strong>{report.listsImported}</strong><span>Lists</span></div>
            <div><strong>{report.hyperlinksImported}</strong><span>Links</span></div>
            <div><strong>{report.sourceMappings}</strong><span>Mappings</span></div>
            <div><strong>{report.footnotesDetected + report.endnotesDetected}</strong><span>Notes</span></div>
            <div><strong>{report.commentsImported}</strong><span>Comments</span></div>
            <div><strong>{report.trackedChangesDetected}</strong><span>Changes</span></div>
            <div><strong>{report.sectionNumberingDetected}</strong><span>Numbering</span></div>
          </div>
          <div className="import-review-summary">
            <span>Parse {Math.round(report.parseTimeMs)}ms</span>
            <span>Render {Math.round(report.renderTimeMs)}ms</span>
            <span>Total {Math.round(report.totalTimeMs)}ms</span>
            <span>{report.trackedChangesMode}</span>
            <span>Rev {report.revision.sourceHash.slice(0, 8)}</span>
          </div>
          <div className="import-review-summary">
            <span>Notes placed {report.notesPlacedExactly}</span>
            <span>Fallback notes {report.notesUsingFallback}</span>
            <span>Broken note refs {report.brokenNoteReferences}</span>
            <span>Internal link fallbacks {report.internalLinkFallbacks}</span>
          </div>
          {report.trackedChangeSummary.length > 0 ? (
            <details className="publishing-section">
              <summary>Tracked-change review summary</summary>
              <div className="validation-list">
                {report.trackedChangeSummary.slice(0, 12).map((change) => (
                  <div className="validation-row" key={change.id}>
                    <span><strong>{change.type}</strong>{change.author ? ` by ${change.author}` : ""}: {change.text || "Formatting change"}</span>
                  </div>
                ))}
                {report.trackedChangeSummary.length > 12 ? <p className="dialog-note">{report.trackedChangeSummary.length - 12} additional tracked changes are included in the JSON report.</p> : null}
              </div>
            </details>
          ) : null}
          {report.fontSubstitutions.length > 0 ? (
            <div className="validation-list">
              <strong>Font substitutions</strong>
              {report.fontSubstitutions.map((item, index) => <div className="validation-row warning" key={`font-${index}`}><AlertTriangle size={15} /><span>{item}</span></div>)}
            </div>
          ) : null}
          {report.layoutSimplifications.length > 0 ? (
            <div className="validation-list">
              <strong>Layout simplifications</strong>
              {report.layoutSimplifications.map((item, index) => <div className="validation-row warning" key={`layout-${index}`}><AlertTriangle size={15} /><span>{item}</span></div>)}
            </div>
          ) : null}
          {report.warnings.length > 0 ? (
            <div className="validation-list">
              <div className="dialog-actions compact">
                <strong>Import warnings</strong>
                <select value={categoryFilter} onChange={(event) => setCategoryFilter(event.currentTarget.value)} aria-label="Filter import warnings">
                  <option value="all">All warnings</option>
                  {categories.map((category) => <option value={category} key={category}>{category}</option>)}
                </select>
              </div>
              {visibleWarnings.map((warning, index) => (
                <div className="validation-row warning" key={`${warning.code}-${index}`}>
                  <AlertTriangle size={15} />
                  <span>{warning.message}{warning.pageNumber ? ` Page ${warning.pageNumber}.` : ""}{warning.sourceBlockId ? ` Source block ${warning.sourceBlockId}.` : ""}</span>
                </div>
              ))}
            </div>
          ) : (
            <p className="dialog-note">Import completed without warnings.</p>
          )}
          <div className="dialog-actions">
            <button className="button ghost" onClick={copyReport}>Copy report</button>
            <button className="button ghost" onClick={downloadReport}>Download JSON</button>
            <button className="button primary" onClick={onClose}>Done</button>
          </div>
        </div>
      </section>
    </div>
  );
}

const publishingZoneOptions: Array<{ value: PublishingZone; label: string }> = [
  { value: "headerLeft", label: "Header left" },
  { value: "headerCenter", label: "Header centre" },
  { value: "headerRight", label: "Header right" },
  { value: "footerLeft", label: "Footer left" },
  { value: "footerCenter", label: "Footer centre" },
  { value: "footerRight", label: "Footer right" },
];

const pageNumberFormatOptions: Array<{ value: PublishingNumberFormat; label: string }> = [
  { value: "decimal", label: "1, 2, 3" },
  { value: "decimal2", label: "01, 02, 03" },
  { value: "decimal3", label: "001, 002, 003" },
  { value: "romanLower", label: "i, ii, iii" },
  { value: "romanUpper", label: "I, II, III" },
  { value: "alphaLower", label: "a, b, c" },
  { value: "alphaUpper", label: "A, B, C" },
  { value: "page", label: "Page 1" },
  { value: "pageOfPages", label: "Page 1 of 12" },
  { value: "custom", label: "Custom template" },
];

const watermarkPositionOptions: Array<{ value: PublishingPositionPreset; label: string }> = [
  { value: "center", label: "Centre" },
  { value: "topLeft", label: "Top left" },
  { value: "topCenter", label: "Top centre" },
  { value: "topRight", label: "Top right" },
  { value: "bottomLeft", label: "Bottom left" },
  { value: "bottomCenter", label: "Bottom centre" },
  { value: "bottomRight", label: "Bottom right" },
  { value: "custom", label: "Custom" },
];

const publishingImageLayoutOptions: Array<{ value: PublishingZoneImageLayout; label: string }> = [
  { value: "imageOnly", label: "Image only" },
  { value: "textBeforeImage", label: "Text before image" },
  { value: "imageBeforeText", label: "Image before text" },
  { value: "imageAboveText", label: "Image above text" },
  { value: "textAboveImage", label: "Text above image" },
];

function PublishingTargetControls({
  target,
  pageCount,
  onChange,
}: {
  target: PublishingSettings["pageNumbers"]["target"];
  pageCount: number;
  onChange: (target: PublishingSettings["pageNumbers"]["target"]) => void;
}) {
  const rangeError = target.mode === "custom" ? getPageRangeError(target.range, pageCount) : "";
  return (
    <div className="publishing-target-controls">
      <label>
        Pages
        <select value={target.mode} onChange={(event) => onChange({ ...target, mode: event.currentTarget.value as PublishingTargetMode })}>
          <option value="all">All pages</option>
          <option value="selected">Selected pages</option>
          <option value="current">Current page</option>
          <option value="custom">Custom range</option>
          <option value="odd">Odd pages</option>
          <option value="even">Even pages</option>
          <option value="exceptFirst">All except first page</option>
          <option value="exceptSelected">All except selected pages</option>
        </select>
      </label>
      {target.mode === "custom" ? (
        <label>
          Page range
          <input value={target.range} onChange={(event) => onChange({ ...target, range: event.currentTarget.value })} placeholder="1-5, 8-10" aria-invalid={Boolean(rangeError)} />
          {rangeError ? <span className="field-error">{rangeError}</span> : null}
        </label>
      ) : null}
      <label className="checkbox-row">
        <input type="checkbox" checked={target.excludeFirst} onChange={(event) => onChange({ ...target, excludeFirst: event.currentTarget.checked })} />
        Exclude first page
      </label>
    </div>
  );
}

function PublishingHeaderFooterZoneControls({
  title,
  zone,
  imageAssets,
  onTextChange,
  onImageChange,
  onRemoveImage,
  onImportImage,
}: {
  title: string;
  zone: HeaderFooterZone;
  imageAssets: Array<{ id: string; name: string; dataUrl?: string; mimeType?: string }>;
  onTextChange: (text: string) => void;
  onImageChange: (patch: Partial<HeaderFooterZoneImage>) => void;
  onRemoveImage: () => void;
  onImportImage: () => void;
}) {
  const image = zone.image;
  const selectedAssetId = image?.assetId ?? "";
  const hasMissingAsset = Boolean(selectedAssetId && !imageAssets.some((asset) => asset.id === selectedAssetId));
  return (
    <details className="publishing-zone-control">
      <summary>{title}</summary>
      <label>
        Text
        <input value={zone.text} onChange={(event) => onTextChange(event.currentTarget.value)} placeholder="{project}, {date}, custom text..." />
      </label>
      <label>
        Image/logo
        <select
          value={selectedAssetId}
          onChange={(event) => {
            const assetId = event.currentTarget.value || undefined;
            onImageChange({ assetId });
          }}
        >
          <option value="">No image</option>
          {hasMissingAsset ? <option value={selectedAssetId}>Missing asset: {selectedAssetId}</option> : null}
          {imageAssets.map((asset) => (
            <option key={asset.id} value={asset.id}>{asset.name}</option>
          ))}
        </select>
      </label>
      <div className="publishing-zone-actions">
        <button type="button" className="link-button" onClick={onImportImage}>Import image</button>
        <button type="button" className="link-button" onClick={onRemoveImage} disabled={!image}>Remove image</button>
      </div>
      {hasMissingAsset ? <p className="field-error" role="alert">This image asset is missing. Replace or remove it before export.</p> : null}
      {image?.assetId ? (
        <div className="publishing-zone-image-options">
          <label>Layout<select value={image.layout} onChange={(event) => onImageChange({ layout: event.currentTarget.value as PublishingZoneImageLayout })}>{publishingImageLayoutOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}</select></label>
          <div className="page-field-grid">
            <label>Width<input type="number" min="1" value={image.width} onChange={(event) => onImageChange({ width: Math.max(1, Number(event.currentTarget.value) || 1) })} /></label>
            <label>Height<input type="number" min="1" value={image.height} onChange={(event) => onImageChange({ height: Math.max(1, Number(event.currentTarget.value) || 1) })} /></label>
          </div>
          <div className="page-field-grid">
            <label>Max width<input type="number" min="1" value={image.maxWidth} onChange={(event) => onImageChange({ maxWidth: Math.max(1, Number(event.currentTarget.value) || 1) })} /></label>
            <label>Max height<input type="number" min="1" value={image.maxHeight} onChange={(event) => onImageChange({ maxHeight: Math.max(1, Number(event.currentTarget.value) || 1) })} /></label>
          </div>
          <div className="page-field-grid">
            <label>Opacity<input type="number" min="0.05" max="1" step="0.05" value={image.opacity} onChange={(event) => onImageChange({ opacity: clamp(Number(event.currentTarget.value), 0.05, 1) })} /></label>
            <label>Padding<input type="number" min="0" value={image.padding} onChange={(event) => onImageChange({ padding: Math.max(0, Number(event.currentTarget.value) || 0) })} /></label>
          </div>
          <div className="page-field-grid">
            <label>Offset X<input type="number" value={image.offsetX} onChange={(event) => onImageChange({ offsetX: Number(event.currentTarget.value) || 0 })} /></label>
            <label>Offset Y<input type="number" value={image.offsetY} onChange={(event) => onImageChange({ offsetY: Number(event.currentTarget.value) || 0 })} /></label>
          </div>
          <div className="page-field-grid">
            <label>Horizontal<select value={image.horizontalAlign} onChange={(event) => onImageChange({ horizontalAlign: event.currentTarget.value as HeaderFooterZoneImage["horizontalAlign"] })}><option value="left">Left</option><option value="center">Centre</option><option value="right">Right</option></select></label>
            <label>Vertical<select value={image.verticalAlign} onChange={(event) => onImageChange({ verticalAlign: event.currentTarget.value as HeaderFooterZoneImage["verticalAlign"] })}><option value="top">Top</option><option value="middle">Middle</option><option value="bottom">Bottom</option></select></label>
          </div>
          <label className="checkbox-row"><input type="checkbox" checked={image.maintainAspectRatio} onChange={(event) => onImageChange({ maintainAspectRatio: event.currentTarget.checked })} />Maintain aspect ratio</label>
        </div>
      ) : null}
    </details>
  );
}

function PublishingAdvancedHeaderFooterControls({
  settings,
  imageAssets,
  onTextChange,
  onImageChange,
  onRemoveImage,
  onImportImage,
}: {
  settings: PublishingSettings;
  imageAssets: Array<{ id: string; name: string; dataUrl?: string; mimeType?: string }>;
  onTextChange: (key: "firstPage" | "oddPage" | "evenPage", zone: PublishingZone, value: string) => void;
  onImageChange: (key: "firstPage" | "oddPage" | "evenPage", zone: PublishingZone, patch: Partial<HeaderFooterZoneImage>) => void;
  onRemoveImage: (key: "firstPage" | "oddPage" | "evenPage", zone: PublishingZone) => void;
  onImportImage: (key: "firstPage" | "oddPage" | "evenPage", zone: PublishingZone) => void;
}) {
  const groups: Array<{ key: "firstPage" | "oddPage" | "evenPage"; title: string }> = [
    { key: "firstPage", title: "First page overrides" },
    { key: "oddPage", title: "Odd page overrides" },
    { key: "evenPage", title: "Even page overrides" },
  ];
  return (
    <div className="publishing-advanced-groups">
      {groups.map((group) => (
        <details key={group.key} className="publishing-advanced-group">
          <summary>{group.title}</summary>
          {publishingZoneOptions.map((zone) => (
            <PublishingHeaderFooterZoneControls
              key={`${group.key}-${zone.value}`}
              title={zone.label}
              zone={mergeHeaderFooterZone(settings.headerFooter[group.key]?.[zone.value])}
              imageAssets={imageAssets}
              onTextChange={(text) => onTextChange(group.key, zone.value, text)}
              onImageChange={(patch) => onImageChange(group.key, zone.value, patch)}
              onRemoveImage={() => onRemoveImage(group.key, zone.value)}
              onImportImage={() => onImportImage(group.key, zone.value)}
            />
          ))}
        </details>
      ))}
    </div>
  );
}

function PublishingStyleControls({
  style,
  onChange,
}: {
  style: PublishingSettings["pageNumbers"]["style"];
  onChange: (style: Partial<PublishingSettings["pageNumbers"]["style"]>) => void;
}) {
  return (
    <div className="publishing-style-controls">
      <label>Font<select value={style.fontFamily} onChange={(event) => onChange({ fontFamily: event.currentTarget.value as PublishingSettings["pageNumbers"]["style"]["fontFamily"] })}><option value="Helvetica">Helvetica</option><option value="Times Roman">Times Roman</option><option value="Courier">Courier</option></select></label>
      <div className="page-field-grid">
        <label>Size<input type="number" min="6" max="144" value={style.fontSize} onChange={(event) => onChange({ fontSize: Number(event.currentTarget.value) || 10 })} /></label>
        <label>Opacity<input type="number" min="0.05" max="1" step="0.05" value={style.opacity} onChange={(event) => onChange({ opacity: clamp(Number(event.currentTarget.value), 0.05, 1) })} /></label>
      </div>
      <label>Colour<input type="color" value={style.color} onChange={(event) => onChange({ color: event.currentTarget.value })} /></label>
      <div className="segmented" role="group" aria-label="Font emphasis">
        <button className={style.bold ? "active" : ""} onClick={() => onChange({ bold: !style.bold })} type="button">Bold</button>
        <button className={style.italic ? "active" : ""} onClick={() => onChange({ italic: !style.italic })} type="button">Italic</button>
      </div>
    </div>
  );
}

function WelcomeScreen({
  logoSrc,
  recentProjects,
  autosave,
  onNewProject,
  onOpenProject,
  onOpenPdf,
  onOpenDocx,
  onOpenPptx,
  onOpenRecent,
  onRemoveRecent,
  onClearRecent,
  onRestoreAutosave,
}: WelcomeScreenProps) {
  return (
    <section className="welcome-screen" aria-label="Publish Pro welcome screen">
      <div className="welcome-hero">
        <img src={logoSrc} alt="" aria-hidden="true" />
        <div>
          <span>Publish Pro</span>
          <h1>Document publishing workspace</h1>
          <p>Create a project, open a PDF, or restore recent local work.</p>
        </div>
      </div>
      <div className="welcome-actions" role="group" aria-label="Start options">
        <button className="button primary" onClick={onNewProject}>
          New Project
        </button>
        <button className="button ghost" onClick={onOpenProject}>
          Open Project
        </button>
        <button className="button ghost" onClick={onOpenPdf}>
          Open PDF
        </button>
        <button className="button ghost" onClick={onOpenDocx}>
          Import Word document
        </button>
        <button className="button ghost" onClick={onOpenPptx}>
          Import PowerPoint presentation
        </button>
      </div>
      {autosave ? (
        <button className="autosave-banner" onClick={onRestoreAutosave}>
          <strong>Recover autosave</strong>
          <span>{autosave.manifest.metadata.name} · {formatDateTime(autosave.savedAt)}</span>
        </button>
      ) : null}
      <div className="welcome-drop-zone">
        <Files size={28} />
        <strong>Drop a PDF, DOCX, PPTX, or .pproj file here</strong>
        <span>Files stay local in your browser.</span>
      </div>
      <div className="recent-projects">
        <div className="section-heading">
          <strong>Recent Projects</strong>
          <span>{recentProjects.length} saved locally</span>
        </div>
        {recentProjects.length > 0 ? (
          recentProjects.map((project) => (
            <div className="recent-project-row" key={project.id}>
              <button onClick={() => onOpenRecent(project)}>
                <strong>{project.name}</strong>
                <span>{project.pageCount} pages · {formatDateTime(project.lastOpenedAt)}{project.savedFilename ? ` · ${project.savedFilename}` : ""}{project.autosaveAvailable ? " · autosave available" : ""}</span>
              </button>
              <button className="link-button" onClick={() => onRemoveRecent(project.id)} aria-label={`Remove ${project.name} from recent projects`}>
                Remove
              </button>
            </div>
          ))
        ) : (
          <p className="muted-text">Recent projects appear here after you save or open a `.pproj` file.</p>
        )}
        {recentProjects.length > 0 ? (
          <button className="link-button recent-clear" onClick={onClearRecent}>
            Clear recent-project history
          </button>
        ) : null}
      </div>
    </section>
  );
}

function ToolCard({
  icon,
  title,
  description,
  onClick,
}: {
  icon: ReactElement;
  title: string;
  description: string;
  onClick: () => void;
}) {
  return (
    <button className="tool-card" onClick={onClick} title={title} aria-label={title}>
      {icon}
      <span>
        <strong>{title}</strong>
        <small>{description}</small>
      </span>
    </button>
  );
}

function RailButton({
  icon,
  label,
  active,
  onClick,
}: {
  icon: ReactElement;
  label: string;
  active?: boolean;
  onClick: () => void;
}) {
  return (
    <button className={`rail-button ${active ? "active" : ""}`} onClick={onClick} title={label} aria-label={label} aria-pressed={active}>
      {icon}
    </button>
  );
}

function getLeftPanelTitle(panel: LeftPanel) {
  if (panel === "insert") return "Insert";
  if (panel === "bookmarks") return "Bookmarks";
  if (panel === "comments") return "Comments";
  if (panel === "search") return "Search";
  return "Pages";
}

function getWorkspacePanelTitle(workspace: WorkspaceMode) {
  if (workspace === "import") return "Import";
  if (workspace === "assemble") return "Pages";
  if (workspace === "review") return "Comments";
  return "Publish";
}

function getWorkspacePanelSubtitle(workspace: WorkspaceMode, pageCount: number, selectedPageCount: number, pdfName: string, commentCount: number) {
  if (workspace === "import") return pdfName;
  if (workspace === "assemble") return `${pageCount} pages${selectedPageCount > 1 ? ` - ${selectedPageCount} selected` : ""}`;
  if (workspace === "review") return `${commentCount} comments`;
  if (workspace === "publish") return "Export and output";
  return "Objects and placement";
}

function getWorkspaceInspectorTitle(workspace: WorkspaceMode) {
  if (workspace === "import") return "Project overview";
  if (workspace === "assemble") return "Page properties";
  if (workspace === "review") return "Object properties";
  return "Export settings";
}

function getWorkspaceInspectorSubtitle(workspace: WorkspaceMode) {
  if (workspace === "import") return "Sources, assets, and project state";
  if (workspace === "assemble") return "Page information, rotation, and labels";
  if (workspace === "review") return "Select an object, comment, or bookmark";
  return "Ready to export the current document";
}

function getWorkspaceEmptyInspectorMessage(workspace: WorkspaceMode) {
  if (workspace === "import") return "Import PDFs, Word documents, or PowerPoint presentations into this project.";
  if (workspace === "assemble") return "Select a page thumbnail to inspect page information, rotation, and labels.";
  if (workspace === "review") return "Select an object, comment, or bookmark to inspect.";
  return "Ready to export. Choose export from the toolbar or Publish panel.";
}

function getStoredThemeMode(): ThemeMode {
  if (typeof window === "undefined") return "system";
  const stored = window.localStorage.getItem(THEME_STORAGE_KEY);
  return stored === "light" || stored === "dark" || stored === "system" ? stored : "system";
}

function normalizeWorkspaceMode(value: unknown): WorkspaceMode {
  if (value === "import" || value === "assemble" || value === "review" || value === "publish") return value;
  if (value === "organise") return "assemble";
  if (value === "edit" || value === "annotate") return "review";
  return "import";
}

function getStoredWorkspaceMode(): WorkspaceMode {
  if (typeof window === "undefined") return "import";
  return normalizeWorkspaceMode(window.localStorage.getItem(WORKSPACE_STORAGE_KEY));
}

function resolveThemeMode(mode: ThemeMode): "light" | "dark" {
  if (mode === "dark") return "dark";
  if (mode === "light") return "light";
  if (typeof window === "undefined") return "light";
  return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

function ProgressOverlay({ message, progress }: WorkState) {
  return (
    <div className="progress-overlay" role="status" aria-live="polite">
      <div className="progress-panel">
        <img className="progress-logo" src={BRAND_ICON_SRC} alt="" aria-hidden="true" />
        <strong>{message}</strong>
        <div className="progress-track" aria-label={progress === undefined ? message : `${Math.round(progress)}% complete`}>
          <span style={{ width: `${clamp(progress ?? 100, 0, 100)}%` }} />
        </div>
      </div>
    </div>
  );
}

function appendCommentsSummary(pdf: PDFDocument, font: PDFFont, boldFont: PDFFont, commentMarks: Mark[]) {
  const exportedComments = commentMarks.filter((mark) => mark.kind === "comment" && mark.comment);
  if (exportedComments.length === 0) return;

  let page = pdf.addPage([612, 792]);
  let y = 742;
  const margin = 54;
  const lineHeight = 15;

  function ensureSpace(requiredHeight: number) {
    if (y - requiredHeight >= 52) return;
    page = pdf.addPage([612, 792]);
    y = 742;
  }

  function drawLine(text: string, size = 10, bold = false, color = "#172033") {
    page.drawText(text || " ", {
      x: margin,
      y,
      size,
      font: bold ? boldFont : font,
      color: colorToRgb(color),
    });
    y -= lineHeight;
  }

  page.drawText("Comments Summary", {
    x: margin,
    y,
    size: 18,
    font: boldFont,
    color: colorToRgb("#172033"),
  });
  y -= 28;
  drawLine("Sticky note markers are embedded on their original PDF pages. Full comment text and replies are listed below.", 9, false, "#475569");
  y -= 8;

  for (const [index, mark] of exportedComments.entries()) {
    if (!mark.comment) continue;
    const bodyLines = wrapText(mark.comment.body || "No comment text yet.", 78);
    const replyLines = mark.comment.replies.flatMap((reply) => [
      `${reply.author} (${formatCommentDate(reply.updatedAt)}):`,
      ...wrapText(reply.body, 74).map((line) => `  ${line}`),
    ]);
    ensureSpace(74 + (bodyLines.length + replyLines.length) * lineHeight);

    const status = mark.comment.resolved ? "Resolved" : "Open";
    drawLine(`${index + 1}. Page ${mark.page} - ${status}`, 12, true);
    drawLine(`Author: ${mark.comment.author} · Created: ${formatCommentDate(mark.comment.createdAt)} · Updated: ${formatCommentDate(mark.comment.updatedAt)}`, 8.5, false, "#64748b");
    for (const line of bodyLines) drawLine(line, 10);
    if (mark.comment.replies.length > 0) {
      drawLine("Replies", 9.5, true, "#475569");
      for (const line of replyLines) drawLine(line, 9);
    }
    y -= 10;
  }
}

function drawTextBoxToPdfPage({
  page,
  mark,
  font,
  x,
  y,
  width,
  height,
  fontSize,
  scale,
}: {
  page: PDFPage;
  mark: Mark;
  font: PDFFont;
  x: number;
  y: number;
  width: number;
  height: number;
  fontSize: number;
  scale: number;
}) {
  const style = getTextBoxStyle(mark);
  const rotate = degrees(mark.rotation);
  const padding = Math.max(0, style.padding * scale);
  const borderWidth = Math.max(0, style.borderWidth * scale);
  const letterSpacing = style.letterSpacing * scale;
  const maxTextWidth = Math.max(1, width - padding * 2 - borderWidth * 2);
  const lines = wrapTextForBox(mark.text, maxTextWidth, fontSize, font, letterSpacing);
  const lineAdvance = fontSize * style.lineHeight;

  if (style.backgroundOpacity > 0) {
    page.drawRectangle({
      x,
      y,
      width,
      height,
      color: colorToRgb(style.backgroundColor),
      opacity: style.backgroundOpacity,
      rotate,
    });
  }

  if (borderWidth > 0) {
    page.drawRectangle({
      x,
      y,
      width,
      height,
      borderColor: colorToRgb(style.borderColor),
      borderWidth,
      rotate,
    });
  }

  const firstBaseline = y + height - padding - fontSize;
  const visibleLineCount = Math.max(1, Math.floor((height - padding * 2) / lineAdvance) + 1);

  for (const [index, line] of lines.slice(0, visibleLineCount).entries()) {
    const lineWidth = measureTextWidth(line, fontSize, font, letterSpacing);
    const alignOffset = style.align === "center" ? (maxTextWidth - lineWidth) / 2 : style.align === "right" ? maxTextWidth - lineWidth : 0;
    const lineX = x + padding + Math.max(0, alignOffset);
    const lineY = firstBaseline - index * lineAdvance;

    drawPdfTextLine(page, line || " ", {
      x: lineX,
      y: lineY,
      size: fontSize,
      font,
      color: mark.color,
      opacity: mark.opacity ?? 1,
      rotate,
      letterSpacing,
    });

    if (style.underline && line) {
      const underlineOffset = -fontSize * 0.16;
      const underlineStart = rotatePdfPoint(lineX, lineY + underlineOffset, lineX, lineY, mark.rotation);
      const underlineEnd = rotatePdfPoint(lineX + lineWidth, lineY + underlineOffset, lineX, lineY, mark.rotation);
      page.drawLine({
        start: underlineStart,
        end: underlineEnd,
        thickness: Math.max(0.4, fontSize * 0.06),
        color: colorToRgb(mark.color),
        opacity: mark.opacity ?? 1,
      });
    }
  }
}

function drawShapeToPdfPage({
  page,
  mark,
  font,
  x,
  y,
  width,
  height,
  scale,
}: {
  page: PDFPage;
  mark: Mark;
  font: PDFFont;
  x: number;
  y: number;
  width: number;
  height: number;
  scale: number;
}) {
  const style = mark.shapeStyle;
  if (!style) return;
  const strokeWidth = Math.max(0.2, style.strokeWidth * scale);
  const dashArray = getPdfDashArray(style.dashStyle, strokeWidth);
  const strokeColor = colorToRgb(style.strokeColor);
  const fillColor = colorToRgb(style.fillColor);
  const rotation = degrees(mark.rotation);

  if (style.shapeKind === "line" || style.shapeKind === "arrow" || style.shapeKind === "doubleArrow") {
    const center = { x: x + width / 2, y: y + height / 2 };
    const start = rotatePdfPoint(x, center.y, center.x, center.y, mark.rotation);
    const end = rotatePdfPoint(x + width, center.y, center.x, center.y, mark.rotation);
    page.drawLine({
      start,
      end,
      thickness: strokeWidth,
      color: strokeColor,
      opacity: style.strokeOpacity,
      dashArray,
    });
    if (style.startArrowhead === "arrow") drawPdfArrowhead(page, start, end, strokeColor, strokeWidth, style.strokeOpacity, true);
    if (style.endArrowhead === "arrow") drawPdfArrowhead(page, start, end, strokeColor, strokeWidth, style.strokeOpacity, false);
    return;
  }

  if (style.shapeKind === "ellipse") {
    page.drawEllipse({
      x: x + width / 2,
      y: y + height / 2,
      xScale: Math.max(1, width / 2),
      yScale: Math.max(1, height / 2),
      color: fillColor,
      opacity: style.fillOpacity,
      borderColor: strokeColor,
      borderOpacity: style.strokeOpacity,
      borderWidth: strokeWidth,
      borderDashArray: dashArray,
      rotate: rotation,
    });
    return;
  }

  if (style.shapeKind === "polygon") {
    const path = `M ${width / 2} ${height} L ${width} ${height * 0.62} L ${width * 0.82} 0 L ${width * 0.18} 0 L 0 ${height * 0.62} Z`;
    page.drawSvgPath(path, {
      x,
      y,
      color: fillColor,
      opacity: style.fillOpacity,
      borderColor: strokeColor,
      borderOpacity: style.strokeOpacity,
      borderWidth: strokeWidth,
      borderDashArray: dashArray,
      rotate: rotation,
    });
    return;
  }

  if (style.shapeKind === "cloud") {
    page.drawSvgPath(getCloudPath(width, height), {
      x,
      y,
      color: fillColor,
      opacity: style.fillOpacity,
      borderColor: strokeColor,
      borderOpacity: style.strokeOpacity,
      borderWidth: strokeWidth,
      borderDashArray: dashArray,
      rotate: rotation,
    });
    return;
  }

  if (style.shapeKind === "roundedRectangle") {
    page.drawSvgPath(getRoundedRectPath(width, height, style.cornerRadius * scale), {
      x,
      y,
      color: fillColor,
      opacity: style.fillOpacity,
      borderColor: strokeColor,
      borderOpacity: style.strokeOpacity,
      borderWidth: strokeWidth,
      borderDashArray: dashArray,
      rotate: rotation,
    });
    return;
  }

  if (style.shapeKind === "callout") {
    const leaderEnd = style.leaderEnd ?? { x: mark.width / 2, y: mark.height + 48 };
    const leader = {
      x: x + (leaderEnd.x / mark.width) * width,
      y: y + height - (leaderEnd.y / mark.height) * height,
    };
    page.drawLine({
      start: { x: x + width / 2, y },
      end: leader,
      thickness: strokeWidth,
      color: strokeColor,
      opacity: style.strokeOpacity,
      dashArray,
    });
    page.drawRectangle({
      x,
      y,
      width,
      height,
      color: fillColor,
      opacity: style.fillOpacity,
      borderColor: strokeColor,
      borderOpacity: style.strokeOpacity,
      borderWidth: strokeWidth,
      borderDashArray: dashArray,
      rotate: rotation,
    });
    page.drawText(style.calloutText || " ", {
      x: x + 8 * scale,
      y: y + height / 2 - 5 * scale,
      size: Math.max(8, 12 * scale),
      font,
      color: strokeColor,
      opacity: style.strokeOpacity,
    });
    return;
  }

  page.drawRectangle({
    x,
    y,
    width,
    height,
    color: fillColor,
    opacity: style.fillOpacity,
    borderColor: strokeColor,
    borderOpacity: style.strokeOpacity,
    borderWidth: strokeWidth,
    borderDashArray: dashArray,
    rotate: rotation,
  });
}

function drawPdfArrowhead(page: PDFPage, start: StrokePoint, end: StrokePoint, color: ReturnType<typeof colorToRgb>, strokeWidth: number, opacity: number, atStart: boolean) {
  const tip = atStart ? start : end;
  const tail = atStart ? end : start;
  const angle = Math.atan2(tip.y - tail.y, tip.x - tail.x);
  const length = Math.max(8, strokeWidth * 5);
  const spread = Math.PI / 7;
  const left = {
    x: tip.x - length * Math.cos(angle - spread),
    y: tip.y - length * Math.sin(angle - spread),
  };
  const right = {
    x: tip.x - length * Math.cos(angle + spread),
    y: tip.y - length * Math.sin(angle + spread),
  };
  page.drawLine({ start: tip, end: left, thickness: strokeWidth, color, opacity });
  page.drawLine({ start: tip, end: right, thickness: strokeWidth, color, opacity });
}

function getRoundedRectPath(width: number, height: number, radius: number) {
  const r = clamp(radius, 0, Math.min(width, height) / 2);
  return `M ${r} 0 L ${width - r} 0 Q ${width} 0 ${width} ${r} L ${width} ${height - r} Q ${width} ${height} ${width - r} ${height} L ${r} ${height} Q 0 ${height} 0 ${
    height - r
  } L 0 ${r} Q 0 0 ${r} 0 Z`;
}

function rotatePdfPoint(x: number, y: number, originX: number, originY: number, rotation: number) {
  const radians = degreesToRadians(rotation);
  const dx = x - originX;
  const dy = y - originY;
  return {
    x: originX + dx * Math.cos(radians) - dy * Math.sin(radians),
    y: originY + dx * Math.sin(radians) + dy * Math.cos(radians),
  };
}

function drawPdfTextLine(
  page: PDFPage,
  text: string,
  options: {
    x: number;
    y: number;
    size: number;
    font: PDFFont;
    color: string;
    opacity: number;
    rotate: ReturnType<typeof degrees>;
    letterSpacing: number;
  }
) {
  if (Math.abs(options.letterSpacing) < 0.01 || text.length <= 1) {
    page.drawText(text, {
      x: options.x,
      y: options.y,
      size: options.size,
      font: options.font,
      color: colorToRgb(options.color),
      opacity: options.opacity,
      rotate: options.rotate,
    });
    return;
  }

  let cursorX = options.x;
  for (const char of text) {
    page.drawText(char, {
      x: cursorX,
      y: options.y,
      size: options.size,
      font: options.font,
      color: colorToRgb(options.color),
      opacity: options.opacity,
      rotate: options.rotate,
    });
    cursorX += options.font.widthOfTextAtSize(char, options.size) + options.letterSpacing;
  }
}

function colorToRgb(hex: string) {
  const clean = hex.replace("#", "");
  const bigint = Number.parseInt(clean, 16);
  const red = ((bigint >> 16) & 255) / 255;
  const green = ((bigint >> 8) & 255) / 255;
  const blue = (bigint & 255) / 255;
  return rgb(red, green, blue);
}

function getTextItemView(item: unknown, index: number, viewportTransform: number[]): TextItemView[] {
  if (!isPdfTextItem(item)) return [];

  const transform = multiplyMatrix(viewportTransform, item.transform);
  const fontSize = Math.max(1, Math.hypot(transform[2], transform[3]));
  const width = Math.max(1, item.width * PAGE_SCALE);
  const height = Math.max(1, fontSize);

  return [
    {
      id: `text-${index}`,
      text: item.str,
      x: transform[4],
      y: transform[5] - height,
      width,
      height,
      fontSize,
    },
  ];
}

function isPdfTextItem(item: unknown): item is { str: string; transform: number[]; width: number } {
  return (
    typeof item === "object" &&
    item !== null &&
    "str" in item &&
    "transform" in item &&
    "width" in item &&
    typeof (item as { str: unknown }).str === "string" &&
    Array.isArray((item as { transform: unknown }).transform) &&
    typeof (item as { width: unknown }).width === "number"
  );
}

function multiplyMatrix(left: number[], right: number[]) {
  return [
    left[0] * right[0] + left[2] * right[1],
    left[1] * right[0] + left[3] * right[1],
    left[0] * right[2] + left[2] * right[3],
    left[1] * right[2] + left[3] * right[3],
    left[0] * right[4] + left[2] * right[5] + left[4],
    left[1] * right[4] + left[3] * right[5] + left[5],
  ];
}

function getPageDialogTitle(dialog: PageDialog) {
  if (dialog === "blank") return "Insert Blank Pages";
  if (dialog === "merge") return "Combine PDF Files";
  if (dialog === "replace") return "Replace Page";
  if (dialog === "split") return "Split PDF";
  if (dialog === "extract") return "Extract Pages";
  if (dialog === "labels") return "Page Labels";
  return "Page Management";
}

function safeRangeCount(range: string, pageCount: number) {
  try {
    return parsePageRanges(range || `1-${pageCount}`, pageCount).flat().length;
  } catch {
    return 0;
  }
}

function getMarkupToolLabel(kind: TextMarkupKind) {
  if (kind === "textHighlight") return "Text highlight";
  if (kind === "underline") return "Underline";
  return "Strikethrough";
}

function isShapeTool(tool: Tool): tool is ShapeTool {
  return typeof tool === "string" && tool.startsWith("shape:");
}

function getShapeKindFromTool(tool: ShapeTool): ShapeKind {
  return tool.replace("shape:", "") as ShapeKind;
}

function getCloudPath(width: number, height: number) {
  const w = Math.max(12, width);
  const h = Math.max(12, height);
  return [
    `M ${w * 0.18} ${h * 0.62}`,
    `C ${w * 0.02} ${h * 0.58}, ${w * 0.02} ${h * 0.34}, ${w * 0.22} ${h * 0.35}`,
    `C ${w * 0.18} ${h * 0.12}, ${w * 0.46} ${h * 0.08}, ${w * 0.52} ${h * 0.25}`,
    `C ${w * 0.66} ${h * 0.08}, ${w * 0.9} ${h * 0.2}, ${w * 0.82} ${h * 0.42}`,
    `C ${w * 1.02} ${h * 0.46}, ${w * 0.95} ${h * 0.75}, ${w * 0.75} ${h * 0.68}`,
    `C ${w * 0.67} ${h * 0.9}, ${w * 0.34} ${h * 0.88}, ${w * 0.34} ${h * 0.7}`,
    `C ${w * 0.28} ${h * 0.78}, ${w * 0.16} ${h * 0.75}, ${w * 0.18} ${h * 0.62}`,
    "Z",
  ].join(" ");
}

function dataUrlToBytes(dataUrl: string) {
  const base64 = dataUrl.split(",")[1];
  if (!base64) throw new Error("Invalid image data.");
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function bytesToDataUrl(bytes: Uint8Array, mimeType: "image/png" | "image/jpeg") {
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return `data:${mimeType};base64,${btoa(binary)}`;
}

async function preparePngSignature(file: File): Promise<PreparedImage> {
  if (file.type !== "image/png" && !file.name.toLowerCase().endsWith(".png")) {
    throw new Error("Unsupported signature type.");
  }

  const originalDataUrl = await readFileAsDataUrl(file);
  const signatureBytes = dataUrlToBytes(originalDataUrl);
  if (!isPngBytes(signatureBytes)) {
    throw new Error("Invalid PNG signature.");
  }

  const dimensions = await getImageDimensions(originalDataUrl);
  return {
    dataUrl: originalDataUrl,
    mimeType: "image/png",
    name: file.name,
    width: dimensions.width,
    height: dimensions.height,
  };
}

async function prepareImageAsset(file: File) {
  const supportedTypes = ["image/png", "image/jpeg", "image/svg+xml", "image/webp"];
  if (!supportedTypes.includes(file.type)) {
    throw new Error("Unsupported image type.");
  }

  const originalDataUrl = await readFileAsDataUrl(file);
  const dimensions = await getImageDimensions(originalDataUrl);

  if (file.type === "image/png") {
    return {
      dataUrl: originalDataUrl,
      mimeType: "image/png" as const,
      name: file.name,
      width: dimensions.width,
      height: dimensions.height,
    };
  }

  if (file.type === "image/jpeg") {
    return {
      dataUrl: originalDataUrl,
      mimeType: "image/jpeg" as const,
      name: file.name,
      width: dimensions.width,
      height: dimensions.height,
    };
  }

  const pngDataUrl = await rasterizeImageToPng(originalDataUrl, dimensions.width, dimensions.height);
  return {
    dataUrl: pngDataUrl,
    mimeType: "image/png" as const,
    name: file.name,
    width: dimensions.width,
    height: dimensions.height,
  };
}

function isPngBytes(bytes: Uint8Array) {
  const pngSignature = [137, 80, 78, 71, 13, 10, 26, 10];
  return pngSignature.every((value, index) => bytes[index] === value);
}

function readFileAsDataUrl(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === "string") resolve(reader.result);
      else reject(new Error("Could not read image."));
    };
    reader.onerror = () => reject(new Error("Could not read image."));
    reader.readAsDataURL(file);
  });
}

function getImageDimensions(dataUrl: string) {
  return new Promise<{ width: number; height: number }>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve({ width: image.naturalWidth || image.width, height: image.naturalHeight || image.height });
    image.onerror = () => reject(new Error("Could not load image."));
    image.src = dataUrl;
  });
}

function rasterizeImageToPng(dataUrl: string, width: number, height: number) {
  return new Promise<string>((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext("2d");
      if (!context) {
        reject(new Error("Could not prepare image."));
        return;
      }
      context.clearRect(0, 0, width, height);
      context.drawImage(image, 0, 0, width, height);
      resolve(canvas.toDataURL("image/png"));
    };
    image.onerror = () => reject(new Error("Could not prepare image."));
    image.src = dataUrl;
  });
}

function hexToRgba(hex: string, alpha: number) {
  const clean = hex.replace("#", "");
  const bigint = Number.parseInt(clean, 16);
  const red = (bigint >> 16) & 255;
  const green = (bigint >> 8) & 255;
  const blue = bigint & 255;
  return `rgb(${red} ${green} ${blue} / ${alpha})`;
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}

function getDefaultImageMarkSize(image: Pick<PreparedImage, "width" | "height">, page: Pick<PageView, "width" | "height">, maxWidthRatio = 0.36, maxHeightRatio = 0.28) {
  const maxWidth = page.width * maxWidthRatio;
  const maxHeight = page.height * maxHeightRatio;
  const scale = Math.min(maxWidth / image.width, maxHeight / image.height, 1);
  return {
    width: Math.max(48, image.width * scale),
    height: Math.max(24, image.height * scale),
  };
}

function getImageAspectRatio(mark: Pick<Mark, "width" | "height" | "imageNaturalWidth" | "imageNaturalHeight">) {
  if (mark.imageNaturalWidth && mark.imageNaturalHeight) return mark.imageNaturalWidth / mark.imageNaturalHeight;
  return mark.width / mark.height || 1;
}

function duplicateMark(mark: Mark, page: number) {
  const offset = 18;
  return {
    ...mark,
    id: crypto.randomUUID(),
    page,
    x: mark.x + offset,
    y: mark.y + offset,
    markupRects: mark.markupRects?.map((rect) => ({ ...rect, x: rect.x + offset, y: rect.y + offset })),
    strokePoints: mark.strokePoints?.map((point) => ({ x: point.x + offset, y: point.y + offset })),
  };
}

function cloneComment(comment: CommentData): CommentData {
  const now = new Date().toISOString();
  return {
    ...comment,
    createdAt: now,
    updatedAt: now,
    replies: comment.replies.map((reply) => ({
      ...reply,
      id: crypto.randomUUID(),
      createdAt: now,
      updatedAt: now,
    })),
  };
}

function renumberPages(pages: PageView[]) {
  return pages.map((page, index) => ({ ...page, pageNumber: index + 1 }));
}

function pageToReference(page: PageView): PageReference {
  return { id: page.id, pageNumber: page.pageNumber, label: page.label, width: page.width, height: page.height };
}

function formatBookmarkDestination(bookmark: Pick<DocumentBookmark, "pageId">, pages: PageView[]) {
  const page = pages.find((item) => item.id === bookmark.pageId);
  if (!page) return "Missing page";
  return page.label ? `${page.label} · Page ${page.pageNumber}` : `Page ${page.pageNumber}`;
}

async function importPdfOutlineBookmarks(pdf: PDFDocumentProxy, pages: PageView[]) {
  try {
    const outline = await pdf.getOutline();
    if (!outline?.length) return [];
    const bookmarks: DocumentBookmark[] = [];
    async function visit(items: Awaited<ReturnType<PDFDocumentProxy["getOutline"]>>, parentId?: string) {
      if (!items) return;
      for (const [order, item] of items.entries()) {
        const pageIndex = await getOutlinePageIndex(pdf, item.dest);
        const page = pageIndex === undefined ? undefined : pages[pageIndex];
        const now = new Date().toISOString();
        const bookmark: DocumentBookmark = {
          id: crypto.randomUUID(),
          title: item.title || (page ? `Page ${page.pageNumber}` : "Unsupported destination"),
          pageId: page?.id ?? `unsupported-${crypto.randomUUID()}`,
          parentId,
          order,
          expanded: true,
          createdAt: now,
          updatedAt: now,
        };
        bookmarks.push(bookmark);
        await visit(item.items, bookmark.id);
      }
    }
    await visit(outline);
    return normalizeBookmarks(bookmarks);
  } catch {
    return [];
  }
}

async function getOutlinePageIndex(pdf: PDFDocumentProxy, destination: unknown) {
  if (!destination) return undefined;
  const resolved = typeof destination === "string" ? await pdf.getDestination(destination) : destination;
  if (!Array.isArray(resolved) || !resolved[0]) return undefined;
  try {
    return await pdf.getPageIndex(resolved[0]);
  } catch {
    return undefined;
  }
}

function migrateMarksToPageIds(marks: Mark[], pages: PageView[]) {
  const pageByNumber = new Map(pages.map((page) => [page.pageNumber, page]));
  const validPageIds = new Set(pages.map((page) => page.id));
  return marks.map((mark) => {
    if (mark.pageId && validPageIds.has(mark.pageId)) return mark;
    const page = pageByNumber.get(mark.page);
    return page ? { ...mark, pageId: page.id } : mark;
  });
}

function syncMarksToPages(marks: Mark[], pages: PageView[]) {
  const pageNumberById = new Map(pages.map((page) => [page.id, page.pageNumber]));
  return marks.map((mark) => {
    if (!mark.pageId) return mark;
    const pageNumber = pageNumberById.get(mark.pageId);
    return pageNumber ? { ...mark, page: pageNumber } : mark;
  });
}

function scaleMarksForReplacement(marks: Mark[], oldPage: PageView, nextPage: PageView) {
  const scaleX = nextPage.width / oldPage.width;
  const scaleY = nextPage.height / oldPage.height;
  return marks.map((mark) => {
    if ((mark.pageId ?? (mark.page === oldPage.pageNumber ? oldPage.id : undefined)) !== oldPage.id) return mark;
    return clampMarkToPage(
      {
        ...mark,
        pageId: oldPage.id,
        x: mark.x * scaleX,
        y: mark.y * scaleY,
        width: mark.width * scaleX,
        height: mark.height * scaleY,
        markupRects: mark.markupRects?.map((rect) => ({
          x: rect.x * scaleX,
          y: rect.y * scaleY,
          width: rect.width * scaleX,
          height: rect.height * scaleY,
        })),
        strokePoints: mark.strokePoints?.map((point) => ({ x: point.x * scaleX, y: point.y * scaleY })),
      },
      [nextPage]
    );
  });
}

function countAnnotationsByPageId(marks: Mark[], pages: PageView[]) {
  const pageIdByNumber = new Map(pages.map((page) => [page.pageNumber, page.id]));
  const counts = new Map<string, number>();
  for (const mark of marks) {
    const pageId = mark.pageId ?? pageIdByNumber.get(mark.page);
    if (!pageId) continue;
    counts.set(pageId, (counts.get(pageId) ?? 0) + 1);
  }
  return counts;
}

function parsePageRanges(input: string, pageCount: number) {
  const trimmed = input.trim();
  if (!trimmed) throw new Error("Enter a page range.");
  const seen = new Set<number>();
  const ranges = trimmed.split(",").map((part) => part.trim()).filter(Boolean);
  if (ranges.length === 0) throw new Error("Enter a page range.");
  return ranges.map((range) => {
    const match = range.match(/^(\d+)(?:-(\d+))?$/);
    if (!match) throw new Error(`Invalid page range: ${range}`);
    const start = Number(match[1]);
    const end = Number(match[2] ?? match[1]);
    if (start < 1 || end < 1 || start > pageCount || end > pageCount || start > end) throw new Error(`Page range is out of bounds: ${range}`);
    const values: number[] = [];
    for (let page = start; page <= end; page += 1) {
      if (seen.has(page)) throw new Error(`Overlapping page range includes page ${page}.`);
      seen.add(page);
      values.push(page);
    }
    return values;
  });
}

function formatSequenceValue(value: number, format: "number" | "roman" | "alpha", padding: number) {
  if (format === "roman") return toRoman(value).toLowerCase();
  if (format === "alpha") return toAlpha(value);
  return String(value).padStart(Math.max(0, padding), "0");
}

function toRoman(value: number) {
  const pairs: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];
  let remaining = Math.max(1, Math.floor(value));
  let output = "";
  for (const [amount, symbol] of pairs) {
    while (remaining >= amount) {
      output += symbol;
      remaining -= amount;
    }
  }
  return output;
}

function toAlpha(value: number) {
  let remaining = Math.max(1, Math.floor(value));
  let output = "";
  while (remaining > 0) {
    remaining -= 1;
    output = String.fromCharCode(65 + (remaining % 26)) + output;
    remaining = Math.floor(remaining / 26);
  }
  return output;
}

function dimensionToPoints(value: number, unit: DimensionUnit) {
  if (unit === "in") return value * 72;
  if (unit === "mm") return (value / 25.4) * 72;
  if (unit === "cm") return (value / 2.54) * 72;
  return value;
}

function remapMarksForPages(marks: Mark[], oldPages: PageView[], nextPages: PageView[]) {
  const pageIdByOldNumber = new Map(oldPages.map((page) => [page.pageNumber, page.id]));
  const pageNumberById = new Map(nextPages.map((page) => [page.id, page.pageNumber]));
  return marks
    .map((mark) => {
      const oldPageId = pageIdByOldNumber.get(mark.page);
      if (!oldPageId) return null;
      const nextPageNumber = pageNumberById.get(oldPageId);
      if (!nextPageNumber) return null;
      return { ...mark, page: nextPageNumber };
    })
    .filter((mark): mark is Mark => Boolean(mark));
}

function createBlankPageView(pageNumber: number, size: { width: number; height: number }, label?: string): PageView {
  return {
    id: crypto.randomUUID(),
    pageNumber,
    width: size.width * PAGE_SCALE,
    height: size.height * PAGE_SCALE,
    rotation: 0,
    label,
    isBlank: true,
    background: "#ffffff",
    imageUrl: "",
    textItems: [],
  };
}

function getBlankPageSize(preset: BlankPagePreset) {
  const sizes: Record<Exclude<BlankPagePreset, "matchCurrent" | "custom">, { width: number; height: number }> = {
    a4Portrait: { width: 595.28, height: 841.89 },
    a4Landscape: { width: 841.89, height: 595.28 },
    a3Portrait: { width: 841.89, height: 1190.55 },
    a3Landscape: { width: 1190.55, height: 841.89 },
    letterPortrait: { width: 612, height: 792 },
    letterLandscape: { width: 792, height: 612 },
    legalPortrait: { width: 612, height: 1008 },
    legalLandscape: { width: 1008, height: 612 },
  };
  if (preset === "custom") return { width: 612, height: 792 };
  if (preset === "matchCurrent") return { width: 612, height: 792 };
  return sizes[preset];
}

async function rotatePageView(page: PageView, delta: 90 | -90 | 180): Promise<PageView> {
  const rotated = {
    ...page,
    width: delta === 180 ? page.width : page.height,
    height: delta === 180 ? page.height : page.width,
    rotation: normalizePageRotation(page.rotation + delta),
    textItems: page.textItems.map((item) => rotateRectForPage(item, page, delta)),
  };
  if (!page.imageUrl) return rotated;
  return {
    ...rotated,
    imageUrl: await rotateImageDataUrl(page.imageUrl, page.width, page.height, delta),
  };
}

function rotateMarkForPage(mark: Mark, page: PageView, delta: 90 | -90 | 180): Mark {
  const rotated = rotateRectForPage(mark, page, delta);
  const markupRects = mark.markupRects?.map((rect) => rotateRectForPage(rect, page, delta));
  const strokePoints = mark.strokePoints?.map((point) =>
    delta === 180
      ? { x: page.width - point.x, y: page.height - point.y }
      : delta === 90
      ? { x: page.height - point.y, y: point.x }
      : { x: point.y, y: page.width - point.x }
  );
  return {
    ...mark,
    ...rotated,
    rotation: normalizePageRotation(mark.rotation + delta),
    markupRects,
    strokePoints,
  };
}

function rotateRectForPage<T extends { x: number; y: number; width: number; height: number }>(rect: T, page: PageView, delta: 90 | -90 | 180): T {
  if (delta === 180) {
    return { ...rect, x: page.width - (rect.x + rect.width), y: page.height - (rect.y + rect.height) };
  }
  if (delta === 90) {
    return { ...rect, x: page.height - (rect.y + rect.height), y: rect.x, width: rect.height, height: rect.width };
  }
  return { ...rect, x: rect.y, y: page.width - (rect.x + rect.width), width: rect.height, height: rect.width };
}

function rotateImageDataUrl(dataUrl: string, width: number, height: number, delta: 90 | -90 | 180) {
  return new Promise<string>((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(delta === 180 ? width : height);
      canvas.height = Math.round(delta === 180 ? height : width);
      const context = canvas.getContext("2d");
      if (!context) {
        reject(new Error("Could not rotate page image."));
        return;
      }
      if (delta === 180) {
        context.translate(canvas.width, canvas.height);
        context.rotate(Math.PI);
      } else if (delta === 90) {
        context.translate(canvas.width, 0);
        context.rotate(Math.PI / 2);
      } else {
        context.translate(0, canvas.height);
        context.rotate(-Math.PI / 2);
      }
      context.drawImage(image, 0, 0, width, height);
      resolve(canvas.toDataURL("image/png"));
    };
    image.onerror = () => reject(new Error("Could not decode page image."));
    image.src = dataUrl;
  });
}

function normalizePageRotation(rotation: number) {
  return ((rotation % 360) + 360) % 360;
}

function downloadPdfBytes(bytes: Uint8Array, filename: string) {
  downloadBytes(bytes, filename, "application/pdf");
}

function downloadBytes(bytes: Uint8Array, filename: string, mimeType: string) {
  const blob = new Blob([bytes.buffer as ArrayBuffer], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

async function saveProjectBytes(bytes: Uint8Array, filename: string) {
  const picker = (window as Window & { showSaveFilePicker?: (options: unknown) => Promise<FileSystemWritableFileHandle> }).showSaveFilePicker;
  if (picker) {
    try {
      const handle = await picker({
        suggestedName: filename,
        types: [
          {
            description: "Publish Pro Project",
            accept: { "application/vnd.publish-pro.project+zip": [".pproj"] },
          },
        ],
      });
      const writable = await handle.createWritable();
      await writable.write(new Blob([bytes.buffer as ArrayBuffer], { type: "application/vnd.publish-pro.project+zip" }));
      await writable.close();
      return true;
    } catch {
      downloadBytes(bytes, filename, "application/vnd.publish-pro.project+zip");
      return false;
    }
  }

  downloadBytes(bytes, filename, "application/vnd.publish-pro.project+zip");
  return false;
}

type FileSystemWritableFileHandle = {
  createWritable: () => Promise<{
    write: (data: Blob) => Promise<void>;
    close: () => Promise<void>;
  }>;
};

function sanitizeFilename(value: string) {
  return (value.trim() || "publish-pro-project").replace(/[<>:"/\\|?*\x00-\x1f]/g, "-").slice(0, 120);
}

function formatDateTime(value: string) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Unknown";
  return new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short" }).format(date);
}

function getProjectErrorMessage(error: unknown, action: "open" | "save" | "restore") {
  const detail = error instanceof Error ? error.message : "An unknown project error occurred.";
  if (action === "save") return `Could not save the Publish Pro project. ${detail}`;
  if (action === "restore") return `Could not restore the autosaved project. ${detail}`;
  return `Could not open this Publish Pro project. ${detail}`;
}

function getProjectFingerprint(
  projectId: string,
  metadata: ProjectMetadata,
  pages: PageView[],
  marks: Mark[],
  sourceDocuments: Record<string, SourceDocument>,
  pdfName: string,
  publishingSettings: PublishingSettings,
  projectImageAssets: ProjectImageAssetLike[] = [],
  navigation?: NavigationState
) {
  return JSON.stringify({
    projectId,
    metadata,
    pdfName,
    pages: pages.map((page) => ({
      id: page.id,
      pageNumber: page.pageNumber,
      sourceDocumentId: page.sourceDocumentId,
      sourcePageNumber: page.sourcePageNumber,
      width: page.width,
      height: page.height,
      rotation: page.rotation,
      label: page.label,
      isBlank: page.isBlank,
      background: page.background,
      imageUrl: page.imageUrl,
      textItems: page.textItems,
      importMetadata: page.importMetadata,
      sourceMappings: page.sourceMappings,
    })),
    marks,
    navigation,
    projectImageAssets: projectImageAssets.map((asset) => ({
      id: asset.id,
      name: asset.name,
      mimeType: asset.mimeType,
      size: estimateDataUrlSize(asset.dataUrl),
      contentHash: asset.dataUrl,
    })),
    publishingSettings,
    sources: Object.values(sourceDocuments).map((source) => ({
      id: source.id,
      name: source.name,
      mimeType: source.mimeType,
      size: source.bytes.byteLength,
    })),
  });
}

function getPublishingCssFontFamily(fontFamily?: string) {
  if (fontFamily === "Times Roman") return "Times New Roman, Times, serif";
  if (fontFamily === "Courier") return "Courier New, Courier, monospace";
  return "Helvetica, Arial, sans-serif";
}

function getTocStandardFontName(fontFamily: TocSettings["fontFamily"], bold: boolean) {
  if (fontFamily === "Times Roman") return bold ? StandardFonts.TimesRomanBold : StandardFonts.TimesRoman;
  if (fontFamily === "Courier") return bold ? StandardFonts.CourierBold : StandardFonts.Courier;
  return bold ? StandardFonts.HelveticaBold : StandardFonts.Helvetica;
}

function createTocPreviewDataUrl(layout: TocLayoutResult, pageId: string, settings: TocSettings) {
  const page = layout.pages.find((item) => item.id === pageId);
  const escape = (value: string) => value.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[char] ?? char));
  const lines = page?.lines.slice(0, 34) ?? [];
  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${layout.width}" height="${layout.height}" viewBox="0 0 ${layout.width} ${layout.height}">
      <rect width="100%" height="100%" fill="#ffffff"/>
      <text x="${settings.marginLeft}" y="${settings.marginTop}" font-family="${getPublishingCssFontFamily(settings.fontFamily)}" font-size="${settings.fontSize * 1.7}" font-weight="700" fill="${settings.color}">${escape(settings.title)}</text>
      ${lines
        .map((line) => {
          const y = layout.height - line.y;
          const pageText = line.pageText ? `<text x="${layout.width - settings.marginRight}" y="${y}" text-anchor="end" font-family="${getPublishingCssFontFamily(settings.fontFamily)}" font-size="${settings.fontSize}" fill="${settings.color}">${escape(line.pageText)}</text>` : "";
          return `<text x="${line.x}" y="${y}" font-family="${getPublishingCssFontFamily(settings.fontFamily)}" font-size="${settings.fontSize}" font-weight="${settings.boldLevels.includes(line.level) ? 700 : 400}" fill="${settings.color}">${escape(line.text)}</text>${pageText}`;
        })
        .join("")}
    </svg>`;
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

function validatePublishingSettings(settings: PublishingSettings) {
  for (const target of [settings.pageNumbers.target, settings.headerFooter.target, settings.watermark.target]) {
    if (target.mode === "custom" && !target.range.trim()) return "Enter a custom page range or choose another page target.";
  }
  return "";
}

function collectPublishingImageAssetIds(settings: PublishingSettings) {
  const ids: string[] = [];
  if (settings.watermark.type === "image" && settings.watermark.imageAssetId) ids.push(settings.watermark.imageAssetId);
  for (const area of ["header", "footer"] as const) {
    for (const side of ["left", "center", "right"] as const) {
      const assetId = settings.headerFooter[area][side].image?.assetId;
      if (assetId) ids.push(assetId);
    }
  }
  for (const key of ["firstPage", "oddPage", "evenPage"] as const) {
    for (const zone of Object.values(settings.headerFooter[key] ?? {})) {
      if (typeof zone === "object" && zone.image?.assetId) ids.push(zone.image.assetId);
    }
  }
  return ids;
}

function getSignaturePlacement(start: StrokePoint, current: StrokePoint, signature: PreparedImage) {
  const dx = current.x - start.x;
  const dy = current.y - start.y;
  const distance = Math.hypot(dx, dy);
  const aspectRatio = signature.width / signature.height || 1;

  if (distance < 8) {
    return { x: start.x, y: start.y, width: undefined, height: undefined };
  }

  let width = Math.max(24, Math.abs(dx));
  let height = Math.max(12, width / aspectRatio);
  if (Math.abs(dy) > height) {
    height = Math.abs(dy);
    width = height * aspectRatio;
  }

  return {
    x: dx < 0 ? start.x - width : start.x,
    y: dy < 0 ? start.y - height : start.y,
    width,
    height,
  };
}

function scalePlacement(placement: ReturnType<typeof getSignaturePlacement>, zoom: number) {
  return {
    left: placement.x * zoom,
    top: placement.y * zoom,
    width: (placement.width ?? 0) * zoom,
    height: (placement.height ?? 0) * zoom,
  };
}

function getDraftSelectionStyle(start: StrokePoint, current: StrokePoint, zoom: number) {
  const x = Math.min(start.x, current.x);
  const y = Math.min(start.y, current.y);
  return {
    left: x * zoom,
    top: y * zoom,
    width: Math.abs(current.x - start.x) * zoom,
    height: Math.abs(current.y - start.y) * zoom,
  };
}

function getDraftTextBoxPlacement(start: StrokePoint, current: StrokePoint) {
  const width = Math.abs(current.x - start.x);
  const height = Math.abs(current.y - start.y);
  if (width < 8 && height < 8) {
    return {
      x: start.x,
      y: start.y,
      width: undefined,
      height: undefined,
    };
  }

  return {
    x: Math.min(start.x, current.x),
    y: Math.min(start.y, current.y),
    width,
    height,
  };
}

function getTextBoxStyle(mark: Mark) {
  return createDefaultTextBoxStyle(mark.textStyle);
}

function getTextBoxContainerStyle(mark: Mark, zoom: number) {
  if (mark.kind !== "text") return {};
  const style = getTextBoxStyle(mark);
  return {
    borderColor: style.borderWidth > 0 ? style.borderColor : undefined,
    borderWidth: Math.max(0, style.borderWidth * zoom),
    backgroundColor: style.backgroundOpacity > 0 ? hexToRgba(style.backgroundColor, style.backgroundOpacity) : "transparent",
    opacity: mark.opacity ?? 1,
  };
}

function getTextBoxTextStyle(mark: Mark, zoom: number) {
  const style = getTextBoxStyle(mark);
  return {
    padding: style.padding * zoom,
    fontFamily: getCssFontFamily(style.fontFamily),
    fontSize: mark.size * zoom,
    fontWeight: style.bold ? 700 : 400,
    fontStyle: style.italic ? "italic" : "normal",
    textDecoration: style.underline ? "underline" : "none",
    textAlign: style.align,
    lineHeight: style.lineHeight,
    letterSpacing: style.letterSpacing * zoom,
    color: mark.color,
  };
}

function doesTextOverflow(mark: Mark) {
  if (mark.kind !== "text") return false;
  const style = getTextBoxStyle(mark);
  const usableHeight = Math.max(1, mark.height - style.padding * 2);
  const roughLineHeight = mark.size * style.lineHeight;
  const roughCharsPerLine = Math.max(1, Math.floor((mark.width - style.padding * 2) / Math.max(1, mark.size * 0.52 + style.letterSpacing)));
  const estimatedLines = mark.text
    .split(/\r?\n/)
    .reduce((count, line) => count + Math.max(1, Math.ceil(line.length / roughCharsPerLine)), 0);
  return estimatedLines * roughLineHeight > usableHeight + 1;
}

function normalizeHexColor(value: string) {
  const trimmed = value.trim();
  const prefixed = trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
  return /^#[0-9a-f]{6}$/i.test(prefixed) ? prefixed.toLowerCase() : null;
}

function clampMarkToPage(mark: Mark, pages: Pick<PageView, "pageNumber" | "width" | "height">[]) {
  const page = pages.find((item) => item.pageNumber === mark.page);
  if (!page) return mark;
  const width = Math.max(1, mark.width);
  const height = Math.max(1, mark.height);

  return {
    ...mark,
    width,
    height,
    opacity: mark.opacity === undefined ? undefined : clamp(mark.opacity, 0.1, 1),
    rotation: normalizeRotation(mark.rotation),
    x: clamp(mark.x, 0, Math.max(0, page.width - width)),
    y: clamp(mark.y, 0, Math.max(0, page.height - height)),
  };
}

function resizeMarkFromCorner(mark: Mark, corner: "nw" | "ne" | "sw" | "se", dx: number, dy: number, pages: Pick<PageView, "pageNumber" | "width" | "height">[]) {
  if (mark.kind === "text" || mark.kind === "shape" || mark.kind === "image" || (mark.kind === "pngSignature" && mark.lockAspectRatio === false)) {
    const minWidth = mark.kind === "text" ? 64 : 24;
    const minHeight = mark.kind === "text" ? 32 : mark.kind === "shape" ? 8 : 12;
    const nextWidth = Math.max(minWidth, mark.width + (corner.includes("w") ? -dx : dx));
    const nextHeight = Math.max(minHeight, mark.height + (corner.includes("n") ? -dy : dy));
    const nextX = corner.includes("w") ? mark.x + mark.width - nextWidth : mark.x;
    const nextY = corner.includes("n") ? mark.y + mark.height - nextHeight : mark.y;

    return clampMarkToPage(
      {
        ...mark,
        x: nextX,
        y: nextY,
        width: nextWidth,
        height: nextHeight,
      },
      pages
    );
  }

  if (mark.kind === "pngSignature" && mark.lockAspectRatio === false) {
    const nextWidth = Math.max(24, mark.width + (corner.includes("w") ? -dx : dx));
    const nextHeight = Math.max(12, mark.height + (corner.includes("n") ? -dy : dy));
    const nextX = corner.includes("w") ? mark.x + mark.width - nextWidth : mark.x;
    const nextY = corner.includes("n") ? mark.y + mark.height - nextHeight : mark.y;

    return clampMarkToPage(
      {
        ...mark,
        x: nextX,
        y: nextY,
        width: nextWidth,
        height: nextHeight,
      },
      pages
    );
  }

  const aspectRatio = mark.width / mark.height || 1;
  const horizontalDelta = corner.includes("w") ? -dx : dx;
  const verticalDelta = corner.includes("n") ? -dy : dy;
  const dominantDelta = Math.abs(horizontalDelta) > Math.abs(verticalDelta) ? horizontalDelta : verticalDelta * aspectRatio;
  const nextWidth = Math.max(24, mark.width + dominantDelta);
  const nextHeight = nextWidth / aspectRatio;
  const nextX = corner.includes("w") ? mark.x + mark.width - nextWidth : mark.x;
  const nextY = corner.includes("n") ? mark.y + mark.height - nextHeight : mark.y;

  return clampMarkToPage(
    {
      ...mark,
      x: nextX,
      y: nextY,
      width: nextWidth,
      height: nextHeight,
    },
    pages
  );
}

function getDimensionPatch(mark: Mark, dimension: "width" | "height", value: number) {
  const nextValue = Math.max(1, value);
  if (mark.kind !== "pngSignature" || mark.lockAspectRatio === false) {
    return { [dimension]: nextValue };
  }

  const aspectRatio = getImageAspectRatio(mark);
  if (dimension === "width") {
    return { width: nextValue, height: Math.max(1, nextValue / aspectRatio) };
  }

  return { height: nextValue, width: Math.max(1, nextValue * aspectRatio) };
}

function moveMarkBy(mark: Mark, dx: number, dy: number, pages: Pick<PageView, "pageNumber" | "width" | "height">[]) {
  if (isTextMarkupKind(mark.kind) && mark.markupRects) {
    const page = pages.find((item) => item.pageNumber === mark.page);
    if (!page) return mark;
    const movedRects = moveMarkupRects(mark.markupRects, dx, dy, page);
    const bounds = getMarkupBounds(movedRects ?? []);
    return bounds ? { ...mark, ...bounds, markupRects: movedRects } : mark;
  }

  if (mark.kind !== "stroke" || !mark.strokePoints) {
    return clampMarkToPage({ ...mark, x: mark.x + dx, y: mark.y + dy }, pages);
  }

  const movedPoints = mark.strokePoints.map((point) => ({ x: point.x + dx, y: point.y + dy }));
  const bounds = getStrokeBounds(movedPoints, mark.size);
  if (!bounds) return mark;

  const page = pages.find((item) => item.pageNumber === mark.page);
  if (!page) return { ...mark, strokePoints: movedPoints, ...bounds };

  const correctionX = bounds.x < 0 ? -bounds.x : bounds.x + bounds.width > page.width ? page.width - (bounds.x + bounds.width) : 0;
  const correctionY = bounds.y < 0 ? -bounds.y : bounds.y + bounds.height > page.height ? page.height - (bounds.y + bounds.height) : 0;
  const correctedPoints = movedPoints.map((point) => ({ x: point.x + correctionX, y: point.y + correctionY }));
  const correctedBounds = getStrokeBounds(correctedPoints, mark.size);

  return {
    ...mark,
    ...(correctedBounds ?? bounds),
    strokePoints: correctedPoints,
  };
}

function moveCalloutLeaderBy(mark: Mark, dx: number, dy: number, pages: Pick<PageView, "pageNumber" | "width" | "height">[]) {
  if (mark.kind !== "shape" || mark.shapeStyle?.shapeKind !== "callout") return mark;
  const page = pages.find((item) => item.pageNumber === mark.page);
  const existing = mark.shapeStyle.leaderEnd ?? { x: mark.width / 2, y: mark.height + 48 };
  const maxX = page ? Math.max(mark.width, page.width - mark.x) : mark.width * 2;
  const maxY = page ? Math.max(mark.height, page.height - mark.y) : mark.height * 2;
  return {
    ...mark,
    shapeStyle: {
      ...mark.shapeStyle,
      leaderEnd: {
        x: clamp(existing.x + dx, 0, maxX),
        y: clamp(existing.y + dy, 0, maxY),
      },
    },
  };
}

function smoothStrokePoints(points: StrokePoint[]) {
  if (points.length <= 2) return points;

  const smoothed: StrokePoint[] = [points[0]];
  for (let index = 1; index < points.length - 1; index += 1) {
    const previous = points[index - 1];
    const current = points[index];
    const next = points[index + 1];
    smoothed.push({
      x: (previous.x + current.x * 2 + next.x) / 4,
      y: (previous.y + current.y * 2 + next.y) / 4,
    });
  }
  smoothed.push(points[points.length - 1]);
  return smoothed;
}

function pointsToSvgPath(points: StrokePoint[]) {
  if (points.length === 0) return "";
  if (points.length === 1) return `M ${points[0].x} ${points[0].y} L ${points[0].x + 0.01} ${points[0].y + 0.01}`;

  let path = `M ${points[0].x} ${points[0].y}`;
  for (let index = 1; index < points.length - 1; index += 1) {
    const current = points[index];
    const next = points[index + 1];
    const midX = (current.x + next.x) / 2;
    const midY = (current.y + next.y) / 2;
    path += ` Q ${current.x} ${current.y} ${midX} ${midY}`;
  }

  const last = points[points.length - 1];
  path += ` L ${last.x} ${last.y}`;
  return path;
}

function pointsRelativeToMark(points: StrokePoint[], mark: Mark) {
  return points.map((point) => ({ x: point.x - mark.x, y: point.y - mark.y }));
}

function getStrokeBounds(points: StrokePoint[], strokeWidth: number) {
  if (points.length === 0) return null;

  const padding = Math.max(2, strokeWidth / 2 + 2);
  const minX = Math.min(...points.map((point) => point.x));
  const minY = Math.min(...points.map((point) => point.y));
  const maxX = Math.max(...points.map((point) => point.x));
  const maxY = Math.max(...points.map((point) => point.y));

  return {
    x: minX - padding,
    y: minY - padding,
    width: Math.max(strokeWidth + padding * 2, maxX - minX + padding * 2),
    height: Math.max(strokeWidth + padding * 2, maxY - minY + padding * 2),
  };
}

function areMarksEqual(left: Mark[], right: Mark[]) {
  if (left.length !== right.length) return false;

  return left.every((mark, index) => {
    const other = right[index];
    return (
      other !== undefined &&
      mark.id === other.id &&
      mark.kind === other.kind &&
      mark.page === other.page &&
      mark.x === other.x &&
      mark.y === other.y &&
      mark.width === other.width &&
      mark.height === other.height &&
      mark.text === other.text &&
      mark.color === other.color &&
      mark.size === other.size &&
      mark.rotation === other.rotation &&
      mark.imageDataUrl === other.imageDataUrl &&
      mark.imageMimeType === other.imageMimeType &&
      mark.imageName === other.imageName &&
      mark.imageNaturalWidth === other.imageNaturalWidth &&
      mark.imageNaturalHeight === other.imageNaturalHeight &&
      mark.opacity === other.opacity &&
      mark.thickness === other.thickness &&
      mark.lockAspectRatio === other.lockAspectRatio &&
      JSON.stringify(mark.comment ?? null) === JSON.stringify(other.comment ?? null) &&
      JSON.stringify(mark.textStyle ?? null) === JSON.stringify(other.textStyle ?? null) &&
      JSON.stringify(mark.shapeStyle ?? null) === JSON.stringify(other.shapeStyle ?? null) &&
      JSON.stringify(mark.markupRects ?? []) === JSON.stringify(other.markupRects ?? []) &&
      mark.strokeOpacity === other.strokeOpacity &&
      JSON.stringify(mark.strokePoints ?? []) === JSON.stringify(other.strokePoints ?? [])
    );
  });
}

function arePagesEqual(left: PageView[], right: PageView[]) {
  if (left.length !== right.length) return false;

  return left.every((page, index) => {
    const other = right[index];
    return (
      other !== undefined &&
      page.id === other.id &&
      page.pageNumber === other.pageNumber &&
      page.sourceDocumentId === other.sourceDocumentId &&
      page.sourcePageNumber === other.sourcePageNumber &&
      page.width === other.width &&
      page.height === other.height &&
      page.rotation === other.rotation &&
      page.label === other.label &&
      page.isBlank === other.isBlank &&
      page.background === other.background &&
      page.imageUrl === other.imageUrl &&
      JSON.stringify(page.textItems) === JSON.stringify(other.textItems)
    );
  });
}

function arraysEqual(left: string[], right: string[]) {
  if (left.length !== right.length) return false;
  return left.every((item, index) => item === right[index]);
}

function getMarkLabel(mark: Mark) {
  if (mark.kind === "shape" && mark.shapeStyle) {
    return `${shapeToolLabels[mark.shapeStyle.shapeKind]} annotation`;
  }
  if (mark.kind === "comment" && mark.comment) {
    return `${mark.comment.resolved ? "Resolved" : "Open"} comment annotation: ${getCommentPreview(mark.comment.body)}`;
  }
  if (mark.kind === "stroke") return "Freehand stroke annotation";
  if (mark.kind === "textHighlight") return "Text highlight annotation";
  if (mark.kind === "underline") return "Underline annotation";
  if (mark.kind === "strikethrough") return "Strikethrough annotation";
  if (mark.kind === "pngSignature") return `PNG signature annotation: ${mark.imageName || mark.text || "Signature"}`;
  if (mark.kind === "image") return `Image annotation: ${mark.imageName || mark.text || "Image"}`;
  if (mark.kind === "highlight") return "Highlight annotation";
  if (mark.kind === "signature") return `Signature annotation: ${mark.text}`;
  if (mark.kind === "stamp") return `Stamp annotation: ${mark.text}`;
  return `Text annotation: ${mark.text}`;
}

function getPdfErrorMessage(error: unknown, phase: "upload" | "render" | "export") {
  const raw = error instanceof Error ? `${error.name} ${error.message}` : String(error);
  const details = raw.toLowerCase();

  if (details.includes("password") || details.includes("encrypted")) {
    return "This PDF is encrypted or password protected. Please unlock it and upload it again.";
  }

  if (details.includes("invalid") || details.includes("corrupt") || details.includes("damaged")) {
    return "This PDF appears to be corrupted or invalid. Please try a different PDF file.";
  }

  if (phase === "render") {
    return "The PDF opened, but one or more pages could not be rendered. Please try a different PDF or export it again from the source app.";
  }

  if (phase === "export") {
    return "The edited PDF could not be exported. Please check the document and try again.";
  }

  return "The PDF could not be opened. Please choose a valid, unencrypted PDF file.";
}

function getImageErrorMessage(error: unknown) {
  const raw = error instanceof Error ? error.message : String(error);
  if (raw.toLowerCase().includes("unsupported")) {
    return "Please choose a PNG, JPG, SVG, or WebP image.";
  }
  return "The image could not be added. Please choose a valid PNG, JPG, SVG, or WebP file.";
}

function getSignatureErrorMessage(error: unknown) {
  const raw = error instanceof Error ? error.message : String(error);
  if (raw.toLowerCase().includes("unsupported")) {
    return "Please choose a PNG signature file. Other image formats are not supported for signatures.";
  }
  return "The signature could not be added. Please choose a valid PNG file.";
}

function normalizeRotation(value: number) {
  if (!Number.isFinite(value)) return 0;
  return clamp(value, -180, 180);
}

function degreesToRadians(value: number) {
  return (value * Math.PI) / 180;
}

function isEditableElement(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) return false;
  return target.matches("input, textarea, select, [contenteditable='true']");
}
