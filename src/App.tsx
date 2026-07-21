import {
  BookOpen,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
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
import { AppChrome } from "./components/app-shell/AppChrome";

GlobalWorkerOptions.workerSrc = pdfWorker;

type ShapeTool = `shape:${ShapeKind}`;
type Tool = "select" | "text" | "highlight" | "signature" | "stamp" | "image" | "draw" | "pngSignature" | "comment" | ShapeTool | TextMarkupKind;
type MarkKind = "text" | "highlight" | "signature" | "stamp" | "image" | "stroke" | "pngSignature" | "comment" | "shape" | TextMarkupKind;
type LeftPanel = "pages" | "files" | "bookmarks" | "comments" | "search";
type WorkState = {
  message: string;
  progress?: number;
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

type PageView = {
  id: string;
  pageNumber: number;
  sourceDocumentId?: string;
  sourcePageNumber?: number;
  width: number;
  height: number;
  rotation: number;
  label?: string;
  isBlank?: boolean;
  background?: string;
  imageUrl: string;
  textItems: TextItemView[];
};

type SourceDocument = {
  id: string;
  name: string;
  bytes: Uint8Array;
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
  const [pdfName, setPdfName] = useState("Untitled document");
  const [pdfBytes, setPdfBytes] = useState<Uint8Array | null>(null);
  const [pdfDoc, setPdfDoc] = useState<PDFDocumentProxy | null>(null);
  const [activeSourceDocumentId, setActiveSourceDocumentId] = useState<string | null>(null);
  const [sourceDocuments, setSourceDocuments] = useState<Record<string, SourceDocument>>({});
  const [pages, setPages] = useState<PageView[]>(demoPages);
  const [currentPage, setCurrentPage] = useState(1);
  const [marks, setMarks] = useState<Mark[]>(initialMarks);
  const [history, setHistory] = useState<HistoryState>({ past: [], future: [] });
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
  const [isLeftPanelCollapsed, setIsLeftPanelCollapsed] = useState(false);
  const [isRightPanelCollapsed, setIsRightPanelCollapsed] = useState(false);
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
  const importPdfInput = useRef<HTMLInputElement>(null);
  const imageInput = useRef<HTMLInputElement>(null);
  const signatureInput = useRef<HTMLInputElement>(null);
  const commentEditorRef = useRef<HTMLTextAreaElement>(null);

  const selected = marks.find((mark) => mark.id === selectedMark) ?? null;
  const comments = marks.filter((mark) => mark.kind === "comment" && mark.comment);
  const currentPageView = pages.find((page) => page.pageNumber === currentPage) ?? pages[0];
  const selectedPageCount = selectedPageIds.filter((id) => pages.some((page) => page.id === id)).length;
  const activePageIds = selectedPageIds.length > 0 ? selectedPageIds.filter((id) => pages.some((page) => page.id === id)) : currentPageView ? [currentPageView.id] : [];
  const pageAnnotations = useMemo(() => countAnnotationsByPageId(marks, pages), [marks, pages]);
  const visibleComments = comments.filter((mark) => {
    if (!mark.comment) return false;
    return (showAllComments || mark.page === currentPage) && commentMatchesFilter(mark.comment.resolved, commentFilter);
  });
  const isBusy = workState !== null;
  const canUndo = history.past.length > 0 && !isBusy;
  const canRedo = history.future.length > 0 && !isBusy;
  const filteredPages = useMemo(() => {
    if (!query.trim()) return pages;
    const normalized = query.trim().toLowerCase();
    return pages.filter((page) => String(page.pageNumber).includes(normalized));
  }, [pages, query]);

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

  useEffect(() => {
    function handleKeyboardShortcut(event: KeyboardEvent) {
      const isModifierPressed = event.metaKey || event.ctrlKey;
      if (!isModifierPressed || isBusy) return;

      const key = event.key.toLowerCase();
      if (isEditableElement(event.target) && (key === "z" || key === "y" || key === "c" || key === "v")) return;
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
  }, [history, isBusy, selected, copiedMark, currentPage, marks, pages]);

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

      setWorkState({ message: "Preparing pages", progress: 35 });
      setPdfName(file.name);
      setPdfBytes(bytes);
      setSourceDocuments({ [sourceId]: { id: sourceId, name: file.name, bytes } });
      setActiveSourceDocumentId(sourceId);
      setPdfDoc(loaded);
      setMarks([]);
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
    if (file) void loadFile(file);
    event.target.value = "";
  }

  async function handleImageUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    try {
      setErrorMessage("");
      setWorkState({ message: `Preparing ${file.name}`, progress: 15 });
      const image = await prepareImageAnnotation(file);
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

  function getPageInsertionIndex(position: InsertPosition) {
    if (position === "start") return 0;
    if (position === "end") return pages.length;
    const activeIds = getActivePageIds();
    const indexes = pages.map((page, index) => (activeIds.includes(page.id) ? index : -1)).filter((index) => index >= 0);
    if (indexes.length === 0) return position === "before" ? currentPage - 1 : currentPage;
    return position === "before" ? Math.min(...indexes) : Math.max(...indexes) + 1;
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
        importedSources[sourceId] = { id: sourceId, name: file.name, bytes };
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
        importedSources[item.sourceDocumentId] = { id: item.sourceDocumentId, name: item.name, bytes: item.bytes };
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
      setSourceDocuments((existing) => ({ ...existing, [replaceFile.sourceDocumentId]: { id: replaceFile.sourceDocumentId, name: replaceFile.name, bytes: replaceFile.bytes } }));
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

  function handlePagesPanelDragOver(event: ReactDragEvent<HTMLDivElement>) {
    if (!hasPdfFiles(event.dataTransfer.files)) return;
    event.preventDefault();
  }

  async function handlePagesPanelDrop(event: ReactDragEvent<HTMLDivElement>) {
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

  async function buildPdfFromPages(exportPages: PageView[], exportMarks: Mark[], onProgress?: (message: string, progress: number) => void) {
    const exportedPdf = await PDFDocument.create();
    const loadedSourcePdfs = new Map<string, PDFDocument>();
    async function getSourcePdf(sourceDocumentId: string) {
      const existing = loadedSourcePdfs.get(sourceDocumentId);
      if (existing) return existing;
      const source = sourceDocuments[sourceDocumentId];
      if (!source) return null;
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

    const pageIdToExportIndex = new Map(exportPages.map((page, index) => [page.id, index]));
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

    appendCommentsSummary(exportedPdf, helvetica, helveticaBold, exportMarks.filter((mark) => mark.kind === "comment" && mark.comment));
    return exportedPdf;
  }

  useEffect(() => {
    if (!import.meta.env.DEV) return;
    const validationWindow = window as Window & {
      __publishProExportForValidation?: () => Promise<{ bytes: number[]; pageCount: number; dimensions: { width: number; height: number }[] }>;
    };
    validationWindow.__publishProExportForValidation = async () => {
      const exportedPdf = await buildPdfFromPages(pages, marks);
      const bytes = await exportedPdf.save();
      return {
        bytes: Array.from(bytes),
        pageCount: exportedPdf.getPageCount(),
        dimensions: exportedPdf.getPages().map((page) => page.getSize()),
      };
    };

    return () => {
      delete validationWindow.__publishProExportForValidation;
    };
  }, [marks, pages, sourceDocuments]);

  async function exportPdf() {
    try {
      setErrorMessage("");
      setWorkState({ message: "Preparing export", progress: 10 });
      const exportedPdf = await buildPdfFromPages(pages, marks, (message, progress) => setWorkState({ message, progress }));
      setWorkState({ message: "Saving PDF", progress: 90 });
      downloadPdfBytes(await exportedPdf.save(), pdfName.replace(/\.pdf$/i, "") + "-edited.pdf");
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
        isBusy={isBusy}
        hasUnsavedChanges={history.past.length > 0}
        onOpen={() => fileInput.current?.click()}
        onExport={() => void exportPdf()}
      />
      <input className="hidden-file-input" ref={fileInput} type="file" accept="application/pdf" onChange={handleUpload} aria-label="Choose PDF file" />
      <input className="hidden-file-input" ref={importPdfInput} type="file" accept="application/pdf" multiple onChange={handleImportPdf} aria-label="Choose PDFs to import" />
      <input className="hidden-file-input" ref={imageInput} type="file" accept="image/png,image/jpeg,image/svg+xml,image/webp" onChange={handleImageUpload} aria-label="Choose image file" />
      <input className="hidden-file-input" ref={signatureInput} type="file" accept="image/png" onChange={handleSignatureUpload} aria-label="Choose PNG signature file" />

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

      <nav className="main-toolbar" aria-label="PDF tools">
        <div className="toolbar-group" aria-label="Navigation tools">
          <ToolButton icon={<MousePointer2 />} label="Select tool" active={activeTool === "select"} onClick={() => setActiveTool("select")} disabled={isBusy} />
          <ToolButton icon={<Hand />} label="Hand tool" disabled />
          <ToolButton icon={<ZoomOut />} label="Zoom out" onClick={() => setZoom((value) => Math.max(0.55, value - 0.1))} disabled={isBusy} />
          <span className="zoom-label">{Math.round(zoom * 100)}%</span>
          <ToolButton icon={<ZoomIn />} label="Zoom in" onClick={() => setZoom((value) => Math.min(1.6, value + 0.1))} disabled={isBusy} />
          <ToolButton icon={<Maximize2 />} label="Fit page" onClick={() => setZoom(1)} disabled={isBusy} />
          <button className="toolbar-text-button" onClick={() => setZoom(1.25)} disabled={isBusy} title="Fit width" aria-label="Fit width">
            Fit width
          </button>
        </div>
        <div className="toolbar-group" aria-label="Edit tools">
          <ToolButton icon={<TextCursorInput />} label="Add text" active={activeTool === "text"} onClick={() => setActiveTool("text")} disabled={isBusy} />
          <ToolButton
            icon={<ImageIcon />}
            label="Add image"
            active={activeTool === "image"}
            onClick={() => {
              setActiveTool("image");
              imageInput.current?.click();
            }}
            disabled={isBusy}
          />
          <ToolButton
            icon={<PenLine />}
            label="Upload PNG signature"
            active={activeTool === "pngSignature"}
            onClick={() => {
              setActiveTool("pngSignature");
              signatureInput.current?.click();
            }}
            disabled={isBusy}
          />
          <ToolButton icon={<PenLine />} label="Add signature text" active={activeTool === "signature"} onClick={() => setActiveTool("signature")} disabled={isBusy} />
        </div>
        <div className="toolbar-group" aria-label="Annotation tools">
          <ToolButton icon={<Pencil />} label="Draw freehand" active={activeTool === "draw"} onClick={() => setActiveTool("draw")} disabled={isBusy} />
          <ToolButton icon={<Highlighter />} label="Highlight selected text" active={activeTool === "textHighlight"} onClick={() => setActiveTool("textHighlight")} disabled={isBusy} />
          <ToolButton icon={<Underline />} label="Underline selected text" active={activeTool === "underline"} onClick={() => setActiveTool("underline")} disabled={isBusy} />
          <ToolButton icon={<Strikethrough />} label="Strikethrough selected text" active={activeTool === "strikethrough"} onClick={() => setActiveTool("strikethrough")} disabled={isBusy} />
          <ToolButton icon={<Highlighter />} label="Add area highlight" active={activeTool === "highlight"} onClick={() => setActiveTool("highlight")} disabled={isBusy} />
          <ToolButton icon={<MessageSquare />} label="Add comment" active={activeTool === "comment"} onClick={() => setActiveTool("comment")} disabled={isBusy} />
          <ToolButton icon={<Square />} label="Add rectangle" active={activeTool === "shape:rectangle"} onClick={() => setActiveTool("shape:rectangle")} disabled={isBusy} />
          <ToolButton icon={<Circle />} label="Add ellipse" active={activeTool === "shape:ellipse"} onClick={() => setActiveTool("shape:ellipse")} disabled={isBusy} />
          <ToolButton icon={<Minus />} label="Add line" active={activeTool === "shape:line"} onClick={() => setActiveTool("shape:line")} disabled={isBusy} />
          <ToolButton icon={<CornerUpRight />} label="Add arrow" active={activeTool === "shape:arrow"} onClick={() => setActiveTool("shape:arrow")} disabled={isBusy} />
          <ToolButton icon={<CornerUpRight />} label="Add double-ended arrow" active={activeTool === "shape:doubleArrow"} onClick={() => setActiveTool("shape:doubleArrow")} disabled={isBusy} />
          <ToolButton icon={<Square />} label="Add rounded rectangle" active={activeTool === "shape:roundedRectangle"} onClick={() => setActiveTool("shape:roundedRectangle")} disabled={isBusy} />
          <ToolButton icon={<Square />} label="Add polygon" active={activeTool === "shape:polygon"} onClick={() => setActiveTool("shape:polygon")} disabled={isBusy} />
          <ToolButton icon={<Circle />} label="Add cloud" active={activeTool === "shape:cloud"} onClick={() => setActiveTool("shape:cloud")} disabled={isBusy} />
          <ToolButton icon={<MessageSquare />} label="Add text callout" active={activeTool === "shape:callout"} onClick={() => setActiveTool("shape:callout")} disabled={isBusy} />
          <ToolButton icon={<Stamp />} label="Add approval stamp" active={activeTool === "stamp"} onClick={() => setActiveTool("stamp")} disabled={isBusy} />
        </div>
        <div className="toolbar-group" aria-label="Document actions">
          <ToolButton icon={<Undo2 />} label="Undo" onClick={undo} disabled={!canUndo} />
          <ToolButton icon={<Redo2 />} label="Redo" onClick={redo} disabled={!canRedo} />
          <ToolButton icon={<Trash2 />} label="Delete selected annotation" onClick={removeSelectedMark} disabled={!selectedMark || isBusy} />
          <ToolButton icon={<FilePlus2 />} label="Add a blank page" onClick={addPage} disabled={isBusy} />
          <ToolButton icon={<Download />} label="Export edited PDF" onClick={() => void exportPdf()} disabled={isBusy} />
        </div>
      </nav>

      <section className="workspace">
        <nav className="left-nav" aria-label="Navigation panels">
          <RailButton icon={<Files />} label="Files" active={leftPanel === "files" && !isLeftPanelCollapsed} onClick={() => { setLeftPanel("files"); setIsLeftPanelCollapsed(false); }} />
          <RailButton icon={<FilePlus2 />} label="Pages" active={leftPanel === "pages" && !isLeftPanelCollapsed} onClick={() => { setLeftPanel("pages"); setIsLeftPanelCollapsed(false); }} />
          <RailButton icon={<BookOpen />} label="Bookmarks" active={leftPanel === "bookmarks" && !isLeftPanelCollapsed} onClick={() => { setLeftPanel("bookmarks"); setIsLeftPanelCollapsed(false); }} />
          <RailButton icon={<MessageSquare />} label="Comments" active={leftPanel === "comments" && !isLeftPanelCollapsed} onClick={() => { setLeftPanel("comments"); setIsLeftPanelCollapsed(false); }} />
          <RailButton icon={<Search />} label="Search" active={leftPanel === "search" && !isLeftPanelCollapsed} onClick={() => { setLeftPanel("search"); setIsLeftPanelCollapsed(false); }} />
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
                <h2>{getLeftPanelTitle(leftPanel)}</h2>
                <span>{leftPanel === "pages" ? `${pages.length} pages${selectedPageCount > 1 ? ` - ${selectedPageCount} selected` : ""}` : pdfName}</span>
              </div>
              <button className="panel-icon-button" onClick={() => setIsLeftPanelCollapsed(true)} title="Collapse left panel" aria-label="Collapse left panel">
                <PanelLeftClose size={16} />
              </button>
            </div>

            {leftPanel === "pages" ? (
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
                  <section className="page-management-section" aria-label="Page selection">
                    <div className="page-actions">
                      <button className="button ghost" onClick={selectAllPages} disabled={isBusy} title="Select all pages" aria-label="Select all pages">
                        Select all
                      </button>
                      <button className="button ghost" onClick={clearPageSelection} disabled={isBusy} title="Clear page selection" aria-label="Clear page selection">
                        Clear
                      </button>
                    </div>
                  </section>

                  <section className="page-management-section" aria-label="Add and import pages">
                    <div className="page-actions">
                      <button className="button ghost" onClick={() => setActivePageDialog("blank")} disabled={isBusy} title="Open blank page dialog" aria-label="Open blank page dialog">
                        Blank pages
                      </button>
                      <button className="button ghost" onClick={() => setActivePageDialog("merge")} disabled={isBusy} title="Combine PDF files" aria-label="Combine PDF files">
                        <Files size={15} />
                        Combine
                      </button>
                    </div>
                  </section>

                  <section className="page-management-section" aria-label="Page operations">
                    <div className="page-actions">
                      <button className="button ghost" onClick={duplicatePage} disabled={isBusy} title="Duplicate selected pages" aria-label="Duplicate selected pages">
                        <FilePlus2 size={15} />
                        Duplicate
                      </button>
                      <button className="button ghost" onClick={deleteSelectedPages} disabled={isBusy || pages.length === 0} title="Delete selected pages" aria-label="Delete selected pages">
                        <Trash2 size={15} />
                        Delete
                      </button>
                      <button className="button ghost" onClick={() => void rotateSelectedPages(90)} disabled={isBusy} title="Rotate selected pages clockwise" aria-label="Rotate selected pages clockwise">
                        Rotate CW
                      </button>
                      <button className="button ghost" onClick={() => void rotateSelectedPages(-90)} disabled={isBusy} title="Rotate selected pages counter-clockwise" aria-label="Rotate selected pages counter-clockwise">
                        Rotate CCW
                      </button>
                      <button className="button ghost" onClick={() => void rotateSelectedPages(180)} disabled={isBusy} title="Rotate selected pages 180 degrees" aria-label="Rotate selected pages 180 degrees">
                        180
                      </button>
                      <button className="button ghost" onClick={() => setActivePageDialog("replace")} disabled={isBusy} title="Replace selected page" aria-label="Replace selected page">
                        Replace
                      </button>
                    </div>
                  </section>

                  <section className="page-management-section" aria-label="Page clipboard and export">
                    <div className="page-actions">
                      <button className="button ghost" onClick={copySelectedPages} disabled={isBusy} title="Copy selected pages" aria-label="Copy selected pages">
                        Copy
                      </button>
                      <button className="button ghost" onClick={() => pasteCopiedPages("after")} disabled={isBusy || !copiedPages} title="Paste copied pages after selection" aria-label="Paste copied pages after selection">
                        Paste
                      </button>
                      <button className="button ghost" onClick={() => setActivePageDialog("extract")} disabled={isBusy} title="Extract selected pages as a PDF" aria-label="Extract selected pages as a PDF">
                        Extract
                      </button>
                      <button className="button ghost" onClick={() => setActivePageDialog("split")} disabled={isBusy || pages.length <= 1} title="Open split PDF dialog" aria-label="Open split PDF dialog">
                        Split
                      </button>
                    </div>
                  </section>

                  <section className="page-management-section" aria-label="Page labels">
                    <label>
                      Selected page label
                      <input
                        value={pages.find((page) => selectedPageIds.includes(page.id))?.label ?? ""}
                        onChange={(event) => updateSelectedPageLabel(event.target.value)}
                        placeholder="Cover, Section A, A-01..."
                        disabled={isBusy}
                      />
                    </label>
                    <button className="button ghost" onClick={() => setActivePageDialog("labels")} disabled={isBusy} title="Apply label pattern" aria-label="Apply label pattern">
                      Pattern
                    </button>
                  </section>
                </div>
              </>
            ) : null}

            {leftPanel === "files" ? (
              <div className="panel-empty">
                <Files size={22} />
                <strong>{pdfName}</strong>
                <button className="button ghost" onClick={() => fileInput.current?.click()} disabled={isBusy} title="Open PDF" aria-label="Open PDF">
                  Open PDF
                </button>
              </div>
            ) : null}

            {leftPanel === "bookmarks" ? (
              <div className="panel-empty">
                <BookOpen size={22} />
                <strong>No bookmarks yet</strong>
                <span>Bookmark management will appear here.</span>
              </div>
            ) : null}

            {leftPanel === "comments" ? (
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

            {leftPanel === "search" ? (
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
            {pages
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
              ))}
          </div>
        </section>

        {!isRightPanelCollapsed ? (
        <aside className="inspector">
          <div className="inspector-header">
            <div>
              <h2>Properties</h2>
              <span>{selected ? getMarkLabel(selected) : activeTool === "draw" ? "Draw tool" : "Document"}</span>
            </div>
            <button className="panel-icon-button" onClick={() => setIsRightPanelCollapsed(true)} title="Collapse right panel" aria-label="Collapse right panel">
              <PanelRightClose size={16} />
            </button>
          </div>
          {activeTool === "draw" ? (
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
          {selected?.kind === "comment" && selected.comment ? (
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
          ) : selected?.kind === "text" ? (
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
          ) : selected?.kind === "shape" && selected.shapeStyle ? (
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
          ) : selected ? (
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
          ) : (
            <div className="empty-panel">
              <ImageIcon size={30} />
              <p>Select an annotation or choose a tool to add content to the page.</p>
            </div>
          )}

          <section className="comments-panel" aria-label="Comments">
            <div className="comments-panel-header">
              <h3>Comments</h3>
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
          </section>

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
          {workState?.message ?? `${marks.length} edits`}
        </div>
      </footer>
    </main>
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
  if (panel === "files") return "Files";
  if (panel === "bookmarks") return "Bookmarks";
  if (panel === "comments") return "Comments";
  if (panel === "search") return "Search";
  return "Pages";
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

async function prepareImageAnnotation(file: File) {
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
  const blob = new Blob([bytes.buffer as ArrayBuffer], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
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
