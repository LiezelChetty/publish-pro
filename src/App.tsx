import {
  CheckCircle2,
  Download,
  FilePlus2,
  Highlighter,
  Image as ImageIcon,
  MessageSquare,
  MousePointer2,
  PenLine,
  Pencil,
  Plus,
  Redo2,
  Search,
  Stamp,
  Strikethrough,
  TextCursorInput,
  Trash2,
  Undo2,
  Underline,
  Upload,
  ZoomIn,
  ZoomOut,
} from "lucide-react";
import { degrees, PDFDocument, rgb, StandardFonts, type PDFFont } from "pdf-lib";
import {
  GlobalWorkerOptions,
  getDocument,
  type PDFDocumentProxy,
  type PDFPageProxy,
} from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";
import { ChangeEvent, PointerEvent, ReactElement, useEffect, useMemo, useRef, useState } from "react";
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

GlobalWorkerOptions.workerSrc = pdfWorker;

type Tool = "select" | "text" | "highlight" | "signature" | "stamp" | "image" | "draw" | "pngSignature" | "comment" | TextMarkupKind;
type MarkKind = "text" | "highlight" | "signature" | "stamp" | "image" | "stroke" | "pngSignature" | "comment" | TextMarkupKind;
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
  strokePoints?: StrokePoint[];
  strokeOpacity?: number;
};

type HistoryEntry = {
  before: Mark[];
  after: Mark[];
  selectedBefore: string | null;
  selectedAfter: string | null;
};

type HistoryState = {
  past: HistoryEntry[];
  future: HistoryEntry[];
};

type PageView = {
  pageNumber: number;
  sourcePageNumber?: number;
  width: number;
  height: number;
  imageUrl: string;
  textItems: TextItemView[];
};

const PAGE_SCALE = 1.35;
const MAX_HISTORY_ENTRIES = 100;
const BRAND_RED = "#d8342a";
const BRAND_ICON_SRC = "/brand/publish-pro-icon.svg";
const BRAND_LOGO_SRC = "/brand/publish-pro-logo.svg";
const COMMENT_MARKER_SIZE = 32;

const demoPages: PageView[] = [
  {
    pageNumber: 1,
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
  const [zoom, setZoom] = useState(1);
  const [query, setQuery] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [workState, setWorkState] = useState<WorkState | null>(null);
  const [errorMessage, setErrorMessage] = useState("");
  const fileInput = useRef<HTMLInputElement>(null);
  const imageInput = useRef<HTMLInputElement>(null);
  const signatureInput = useRef<HTMLInputElement>(null);
  const commentEditorRef = useRef<HTMLTextAreaElement>(null);

  const selected = marks.find((mark) => mark.id === selectedMark) ?? null;
  const comments = marks.filter((mark) => mark.kind === "comment" && mark.comment);
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

  useEffect(() => {
    function handleKeyboardShortcut(event: KeyboardEvent) {
      const isModifierPressed = event.metaKey || event.ctrlKey;
      if (!isModifierPressed || isBusy) return;

      const key = event.key.toLowerCase();
      const isUndo = key === "z" && !event.shiftKey;
      const isRedo = (key === "z" && event.shiftKey) || key === "y";
      const isCopy = key === "c" && !event.shiftKey && !isEditableElement(event.target);
      const isPaste = key === "v" && !event.shiftKey && !isEditableElement(event.target);

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
          const rendered = await renderPage(page);
          if (cancelled) return;
          nextPages.push(rendered);
          setWorkState({
            message: `Rendering page ${index} of ${activePdf.numPages}`,
            progress: (index / activePdf.numPages) * 100,
          });
        }
        setPages(nextPages);
        setCurrentPage(1);
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
  }, [pdfDoc]);

  async function renderPage(page: PDFPageProxy): Promise<PageView> {
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
      pageNumber: page.pageNumber,
      sourcePageNumber: page.pageNumber,
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

      setWorkState({ message: "Preparing pages", progress: 35 });
      setPdfName(file.name);
      setPdfBytes(bytes);
      setPdfDoc(loaded);
      setMarks([]);
      setHistory({ past: [], future: [] });
      setSelectedMark(null);
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

  function addPage() {
    const pageNumber = pages.length + 1;
    setPages((existing) => [
      ...existing,
      {
        pageNumber,
        sourcePageNumber: undefined,
        width: 612,
        height: 792,
        imageUrl: "",
        textItems: [],
      },
    ]);
    setCurrentPage(pageNumber);
  }

  function addMark(pageNumber: number, x: number, y: number) {
    if (activeTool === "select") return;
    if (activeTool === "image") {
      imageInput.current?.click();
      return;
    }
    if (activeTool === "pngSignature") {
      signatureInput.current?.click();
      return;
    }
    if (activeTool === "draw") return;
    if (isTextMarkupKind(activeTool)) return;

    if (activeTool === "comment") {
      const comment = createCommentData();
      const mark: Mark = {
        id: crypto.randomUUID(),
        kind: "comment",
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
      kind: activeTool === "text" ? "text" : activeTool,
      page: pageNumber,
      x,
      y,
      width: activeTool === "highlight" ? 220 : activeTool === "signature" ? 190 : 124,
      height: activeTool === "text" ? 34 : activeTool === "highlight" ? 28 : 42,
      text:
        activeTool === "text"
          ? "New text"
          : activeTool === "signature"
            ? "Signature"
            : activeTool === "stamp"
              ? "APPROVED"
              : "",
      color: activeTool === "highlight" ? "#facc15" : activeTool === "stamp" ? BRAND_RED : "#111827",
      size: activeTool === "text" ? 18 : 16,
      rotation: 0,
    };

    commitMarks([...marks, mark], mark.id);
    setActiveTool("select");
  }

  function addTextMarkup(pageNumber: number, rects: MarkupRect[]) {
    if (!isTextMarkupKind(activeTool)) return;
    const bounds = getMarkupBounds(rects);
    if (!bounds) return;

    const mark: Mark = {
      id: crypto.randomUUID(),
      kind: activeTool,
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

  function duplicateSelectedMark() {
    if (!selected) return;
    const duplicate = duplicateMark(selected, currentPage);
    commitMarks([...marks, clampMarkToPage(duplicate, pages)], duplicate.id);
  }

  function pasteCopiedMark() {
    if (!copiedMark) return;
    const pasted = duplicateMark(copiedMark, currentPage);
    commitMarks([...marks, clampMarkToPage(pasted, pages)], pasted.id);
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
    setCurrentPage(mark.page);
    setSelectedMark(mark.id);
  }

  function duplicatePage() {
    const sourcePage = pages.find((page) => page.pageNumber === currentPage);
    if (!sourcePage) return;

    const nextNumber = pages.length + 1;
    setPages((existing) => [...existing, { ...sourcePage, pageNumber: nextNumber }]);
    const nextMarks = [
      ...marks,
      ...marks
        .filter((mark) => mark.page === currentPage)
        .map((mark) => ({ ...mark, id: crypto.randomUUID(), page: nextNumber })),
    ];
    commitMarks(nextMarks, selectedMark);
    setCurrentPage(nextNumber);
  }

  function commitMarks(nextMarks: Mark[], nextSelectedMark: string | null, historyBefore = marks) {
    if (areMarksEqual(historyBefore, nextMarks) && selectedMark === nextSelectedMark) return;

    setHistory((existing) => ({
      past: [...existing.past.slice(-(MAX_HISTORY_ENTRIES - 1)), { before: historyBefore, after: nextMarks, selectedBefore: selectedMark, selectedAfter: nextSelectedMark }],
      future: [],
    }));
    setMarks(nextMarks);
    setSelectedMark(nextSelectedMark);
  }

  function undo() {
    setHistory((existing) => {
      const entry = existing.past[existing.past.length - 1];
      if (!entry) return existing;

      setMarks(entry.before);
      setSelectedMark(entry.selectedBefore);

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

      setMarks(entry.after);
      setSelectedMark(entry.selectedAfter);

      return {
        past: [...existing.past, entry].slice(-MAX_HISTORY_ENTRIES),
        future: existing.future.slice(1),
      };
    });
  }

  async function exportPdf() {
    try {
      setErrorMessage("");
      setWorkState({ message: "Preparing export", progress: 10 });
      const exportedPdf = pdfBytes ? await PDFDocument.load(pdfBytes) : await PDFDocument.create();
      const sourcePdf = pdfBytes ? await PDFDocument.load(pdfBytes) : null;
      const helvetica = await exportedPdf.embedFont(StandardFonts.Helvetica);
      const helveticaBold = await exportedPdf.embedFont(StandardFonts.HelveticaBold);

      while (exportedPdf.getPageCount() < pages.length) {
        const pageView = pages[exportedPdf.getPageCount()];
        if (sourcePdf && pageView.sourcePageNumber) {
          const [copiedPage] = await exportedPdf.copyPages(sourcePdf, [pageView.sourcePageNumber - 1]);
          exportedPdf.addPage(copiedPage);
        } else {
          exportedPdf.addPage([612, 792]);
        }
      }

      for (const [index, mark] of marks.entries()) {
        setWorkState({ message: "Applying annotations", progress: 20 + (index / Math.max(1, marks.length)) * 60 });
        const page = exportedPdf.getPage(mark.page - 1);
        const sourcePage = pages.find((item) => item.pageNumber === mark.page);
        if (!sourcePage) continue;

        const { width, height } = page.getSize();
        const x = (mark.x / sourcePage.width) * width;
        const y = height - ((mark.y + mark.height) / sourcePage.height) * height;
        const markWidth = (mark.width / sourcePage.width) * width;
        const markHeight = (mark.height / sourcePage.height) * height;
        const markSize = Math.max(8, (mark.size / sourcePage.height) * height * 1.2);

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

      appendCommentsSummary(exportedPdf, helvetica, helveticaBold, comments);

      setWorkState({ message: "Saving PDF", progress: 90 });
      const exported = await exportedPdf.save();
      const blob = new Blob([exported.buffer as ArrayBuffer], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = pdfName.replace(/\.pdf$/i, "") + "-edited.pdf";
      anchor.click();
      URL.revokeObjectURL(url);
      setWorkState(null);
    } catch (error) {
      setWorkState(null);
      setErrorMessage(getPdfErrorMessage(error, "export"));
    }
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark" aria-hidden="true">
            <img src={BRAND_ICON_SRC} alt="" />
          </div>
          <div>
            <h1>Publish Pro</h1>
            <p>{pdfName}</p>
          </div>
        </div>

        <div className="topbar-actions">
          <button
            className="button ghost"
            onClick={() => fileInput.current?.click()}
            disabled={isBusy}
            title="Upload a PDF"
            aria-label="Upload a PDF"
          >
            <Upload size={18} />
            Upload
          </button>
          <button className="button primary" onClick={() => void exportPdf()} disabled={isBusy} title="Export edited PDF" aria-label="Export edited PDF">
            <Download size={18} />
            Export PDF
          </button>
          <input ref={fileInput} type="file" accept="application/pdf" onChange={handleUpload} aria-label="Choose PDF file" />
          <input ref={imageInput} type="file" accept="image/png,image/jpeg,image/svg+xml,image/webp" onChange={handleImageUpload} aria-label="Choose image file" />
          <input ref={signatureInput} type="file" accept="image/png" onChange={handleSignatureUpload} aria-label="Choose PNG signature file" />
        </div>
      </header>

      <section className="workspace">
        <aside className="left-rail">
          <div className="searchbox">
            <Search size={16} />
            <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Find page" aria-label="Find page" />
          </div>

          <div className="page-list">
            {filteredPages.map((page) => (
              <button
                className={`thumb ${currentPage === page.pageNumber ? "active" : ""}`}
                key={page.pageNumber}
                onClick={() => setCurrentPage(page.pageNumber)}
                title={`Go to page ${page.pageNumber}`}
                aria-label={`Go to page ${page.pageNumber}`}
                aria-current={currentPage === page.pageNumber ? "page" : undefined}
              >
                <div className="thumb-page">
                  {page.imageUrl ? <img src={page.imageUrl} alt="" /> : <FilePlus2 size={32} />}
                </div>
                <span>Page {page.pageNumber}</span>
              </button>
            ))}
          </div>

          <div className="page-actions">
            <button className="button ghost" onClick={addPage} disabled={isBusy} title="Add a blank page" aria-label="Add a blank page">
              <Plus size={17} />
              Add page
            </button>
            <button className="button ghost" onClick={duplicatePage} disabled={isBusy} title="Duplicate current page" aria-label="Duplicate current page">
              <FilePlus2 size={17} />
              Duplicate
            </button>
          </div>
        </aside>

        <section className="editor">
          <nav className="toolstrip" aria-label="PDF tools">
            <ToolButton icon={<MousePointer2 />} label="Select tool" active={activeTool === "select"} onClick={() => setActiveTool("select")} disabled={isBusy} />
            <ToolButton icon={<TextCursorInput />} label="Add text" active={activeTool === "text"} onClick={() => setActiveTool("text")} disabled={isBusy} />
            <ToolButton icon={<Highlighter />} label="Highlight selected text" active={activeTool === "textHighlight"} onClick={() => setActiveTool("textHighlight")} disabled={isBusy} />
            <ToolButton icon={<Underline />} label="Underline selected text" active={activeTool === "underline"} onClick={() => setActiveTool("underline")} disabled={isBusy} />
            <ToolButton icon={<Strikethrough />} label="Strikethrough selected text" active={activeTool === "strikethrough"} onClick={() => setActiveTool("strikethrough")} disabled={isBusy} />
            <ToolButton icon={<Highlighter />} label="Add area highlight" active={activeTool === "highlight"} onClick={() => setActiveTool("highlight")} disabled={isBusy} />
            <ToolButton icon={<MessageSquare />} label="Add comment" active={activeTool === "comment"} onClick={() => setActiveTool("comment")} disabled={isBusy} />
            <ToolButton icon={<Pencil />} label="Draw freehand" active={activeTool === "draw"} onClick={() => setActiveTool("draw")} disabled={isBusy} />
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
            <ToolButton icon={<Stamp />} label="Add approval stamp" active={activeTool === "stamp"} onClick={() => setActiveTool("stamp")} disabled={isBusy} />
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
            <div className="tool-divider" />
            <ToolButton icon={<Undo2 />} label="Undo" onClick={undo} disabled={!canUndo} />
            <ToolButton icon={<Redo2 />} label="Redo" onClick={redo} disabled={!canRedo} />
            <div className="tool-divider" />
            <ToolButton icon={<Trash2 />} label="Delete selected annotation" onClick={removeSelectedMark} disabled={!selectedMark || isBusy} />
            <div className="tool-divider" />
            <ToolButton icon={<ZoomOut />} label="Zoom out" onClick={() => setZoom((value) => Math.max(0.55, value - 0.1))} disabled={isBusy} />
            <span className="zoom-label">{Math.round(zoom * 100)}%</span>
            <ToolButton icon={<ZoomIn />} label="Zoom in" onClick={() => setZoom((value) => Math.min(1.6, value + 0.1))} disabled={isBusy} />
          </nav>

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
              const file = event.dataTransfer.files[0];
              if (file) void loadFile(file);
            }}
          >
            {workState ? <ProgressOverlay message={workState.message} progress={workState.progress} /> : null}
            {pages
              .filter((page) => page.pageNumber === currentPage)
              .map((page) => (
                <DocumentPage
                  key={page.pageNumber}
                  page={page}
                  marks={marks.filter((mark) => mark.page === page.pageNumber)}
                  zoom={zoom}
                  selectedMark={selectedMark}
                  activeTool={activeTool}
                  penColor={penColor}
                  penWidth={penWidth}
                  penOpacity={penOpacity}
                  pendingSignature={pendingSignature}
                  onSelect={setSelectedMark}
                  onAddMark={addMark}
                  onAddStroke={addStroke}
                  onAddTextMarkup={addTextMarkup}
                  onPlaceSignature={placeSignature}
                  onPreviewMark={previewMark}
                  onCommitMarkChange={commitMarkChange}
                />
              ))}
          </div>
        </section>

        <aside className="inspector">
          <h2>Properties</h2>
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
      </section>
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
  onAddStroke,
  onAddTextMarkup,
  onPlaceSignature,
  onPreviewMark,
  onCommitMarkChange,
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
  onAddStroke: (page: number, points: StrokePoint[]) => void;
  onAddTextMarkup: (page: number, rects: MarkupRect[]) => void;
  onPlaceSignature: (page: number, x: number, y: number, width?: number, height?: number) => void;
  onPreviewMark: (id: string, patch: Partial<Mark>) => void;
  onCommitMarkChange: (beforeMark: Mark, afterMark: Mark) => void;
}) {
  const pageRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef<{
    before: Mark;
    latest: Mark;
    startX: number;
    startY: number;
    mode: "move" | "resize";
    corner?: "nw" | "ne" | "sw" | "se";
  } | null>(null);
  const [draftStroke, setDraftStroke] = useState<StrokePoint[]>([]);
  const [draftSignature, setDraftSignature] = useState<{ start: StrokePoint; current: StrokePoint } | null>(null);
  const [draftTextSelection, setDraftTextSelection] = useState<{ start: StrokePoint; current: StrokePoint } | null>(null);
  const activeDraftStroke = Array.isArray(draftStroke) ? draftStroke : [];

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

    if (rects.length > 0) onAddTextMarkup(page.pageNumber, rects);
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
            const drag = dragRef.current;
            if (!drag) return;
            const activeMark = marks.find((mark) => mark.id === drag.before.id);
            if (!activeMark) return;
            const dx = (event.clientX - drag.startX) / zoom;
            const dy = (event.clientY - drag.startY) / zoom;
            const latest =
              drag.mode === "resize" && drag.corner
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
      {marks.map((mark) => (
        <button
          className={`mark mark-${mark.kind} ${selectedMark === mark.id ? "selected" : ""} ${mark.kind === "comment" && mark.comment?.resolved ? "resolved" : ""}`}
          key={mark.id}
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
          }}
          onPointerDown={(event) => {
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
        >
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
          !isTextMarkupKind(mark.kind)
            ? mark.text
            : ""}
          {(mark.kind === "image" || mark.kind === "pngSignature") && selectedMark === mark.id
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
        </button>
      ))}
    </div>
  );
}

function BlankPage() {
  return (
    <div className="blank-page">
      <img className="blank-logo" src={BRAND_LOGO_SRC} alt="Publish Pro" />
      <FilePlus2 size={34} aria-hidden="true" />
      <span>New PDF page</span>
    </div>
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
  onClick: () => void;
}) {
  return (
    <button className={`tool-button ${active ? "active" : ""}`} onClick={onClick} disabled={disabled} title={label} aria-label={label} aria-pressed={active}>
      {icon}
    </button>
  );
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

function getMarkupToolLabel(kind: TextMarkupKind) {
  if (kind === "textHighlight") return "Text highlight";
  if (kind === "underline") return "Underline";
  return "Strikethrough";
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
      JSON.stringify(mark.markupRects ?? []) === JSON.stringify(other.markupRects ?? []) &&
      mark.strokeOpacity === other.strokeOpacity &&
      JSON.stringify(mark.strokePoints ?? []) === JSON.stringify(other.strokePoints ?? [])
    );
  });
}

function getMarkLabel(mark: Mark) {
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
