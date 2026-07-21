import {
  Download,
  FilePlus2,
  Highlighter,
  Image as ImageIcon,
  MousePointer2,
  PenLine,
  Plus,
  Search,
  Stamp,
  TextCursorInput,
  Trash2,
  Upload,
  ZoomIn,
  ZoomOut,
} from "lucide-react";
import { PDFDocument, rgb, StandardFonts } from "pdf-lib";
import {
  GlobalWorkerOptions,
  getDocument,
  type PDFDocumentProxy,
  type PDFPageProxy,
} from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";
import { ChangeEvent, PointerEvent, ReactElement, useEffect, useMemo, useRef, useState } from "react";

GlobalWorkerOptions.workerSrc = pdfWorker;

type Tool = "select" | "text" | "highlight" | "signature" | "stamp";
type MarkKind = "text" | "highlight" | "signature" | "stamp";
type WorkState = {
  message: string;
  progress?: number;
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
};

type PageView = {
  pageNumber: number;
  sourcePageNumber?: number;
  width: number;
  height: number;
  imageUrl: string;
};

const PAGE_SCALE = 1.35;
const BRAND_RED = "#d8342a";

const demoPages: PageView[] = [
  {
    pageNumber: 1,
    width: 612,
    height: 792,
    imageUrl: "",
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
  },
];

export function App() {
  const [pdfName, setPdfName] = useState("Untitled document");
  const [pdfBytes, setPdfBytes] = useState<Uint8Array | null>(null);
  const [pdfDoc, setPdfDoc] = useState<PDFDocumentProxy | null>(null);
  const [pages, setPages] = useState<PageView[]>(demoPages);
  const [currentPage, setCurrentPage] = useState(1);
  const [marks, setMarks] = useState<Mark[]>(initialMarks);
  const [activeTool, setActiveTool] = useState<Tool>("select");
  const [selectedMark, setSelectedMark] = useState<string | null>("demo-title");
  const [zoom, setZoom] = useState(1);
  const [query, setQuery] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [workState, setWorkState] = useState<WorkState | null>(null);
  const [errorMessage, setErrorMessage] = useState("");
  const fileInput = useRef<HTMLInputElement>(null);

  const selected = marks.find((mark) => mark.id === selectedMark) ?? null;
  const isBusy = workState !== null;
  const filteredPages = useMemo(() => {
    if (!query.trim()) return pages;
    const normalized = query.trim().toLowerCase();
    return pages.filter((page) => String(page.pageNumber).includes(normalized));
  }, [pages, query]);

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

    return {
      pageNumber: page.pageNumber,
      sourcePageNumber: page.pageNumber,
      width: viewport.width,
      height: viewport.height,
      imageUrl: canvas.toDataURL("image/png"),
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
      },
    ]);
    setCurrentPage(pageNumber);
  }

  function addMark(pageNumber: number, x: number, y: number) {
    if (activeTool === "select") return;

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
    };

    setMarks((existing) => [...existing, mark]);
    setSelectedMark(mark.id);
    setActiveTool("select");
  }

  function updateMark(id: string, patch: Partial<Mark>) {
    setMarks((existing) => existing.map((mark) => (mark.id === id ? { ...mark, ...patch } : mark)));
  }

  function removeSelectedMark() {
    if (!selectedMark) return;
    setMarks((existing) => existing.filter((mark) => mark.id !== selectedMark));
    setSelectedMark(null);
  }

  function duplicatePage() {
    const sourcePage = pages.find((page) => page.pageNumber === currentPage);
    if (!sourcePage) return;

    const nextNumber = pages.length + 1;
    setPages((existing) => [...existing, { ...sourcePage, pageNumber: nextNumber }]);
    setMarks((existing) => [
      ...existing,
      ...existing
        .filter((mark) => mark.page === currentPage)
        .map((mark) => ({ ...mark, id: crypto.randomUUID(), page: nextNumber })),
    ]);
    setCurrentPage(nextNumber);
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

      marks.forEach((mark, index) => {
        setWorkState({ message: "Applying annotations", progress: 20 + (index / Math.max(1, marks.length)) * 60 });
        const page = exportedPdf.getPage(mark.page - 1);
        const sourcePage = pages.find((item) => item.pageNumber === mark.page);
        if (!sourcePage) return;

        const { width, height } = page.getSize();
        const x = (mark.x / sourcePage.width) * width;
        const y = height - ((mark.y + mark.height) / sourcePage.height) * height;
        const markWidth = (mark.width / sourcePage.width) * width;
        const markHeight = (mark.height / sourcePage.height) * height;
        const markSize = Math.max(8, (mark.size / sourcePage.height) * height * 1.2);

        if (mark.kind === "highlight") {
          page.drawRectangle({
            x,
            y,
            width: markWidth,
            height: markHeight,
            color: colorToRgb(mark.color),
            opacity: 0.45,
          });
          return;
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
      });

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
          <div className="brand-mark">P</div>
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
            <ToolButton icon={<Highlighter />} label="Add highlight" active={activeTool === "highlight"} onClick={() => setActiveTool("highlight")} disabled={isBusy} />
            <ToolButton icon={<PenLine />} label="Add signature text" active={activeTool === "signature"} onClick={() => setActiveTool("signature")} disabled={isBusy} />
            <ToolButton icon={<Stamp />} label="Add approval stamp" active={activeTool === "stamp"} onClick={() => setActiveTool("stamp")} disabled={isBusy} />
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
                  onSelect={setSelectedMark}
                  onAddMark={addMark}
                  onUpdateMark={updateMark}
                />
              ))}
          </div>
        </section>

        <aside className="inspector">
          <h2>Properties</h2>
          {selected ? (
            <div className="control-stack">
              <label>
                Content
                <textarea
                  value={selected.text}
                  onChange={(event) => updateMark(selected.id, { text: event.target.value })}
                  aria-label="Annotation content"
                />
              </label>
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
              <div className="dimension-grid">
                <label>
                  W
                  <input
                    type="number"
                    value={Math.round(selected.width)}
                    onChange={(event) => updateMark(selected.id, { width: Number(event.target.value) })}
                    aria-label="Annotation width"
                  />
                </label>
                <label>
                  H
                  <input
                    type="number"
                    value={Math.round(selected.height)}
                    onChange={(event) => updateMark(selected.id, { height: Number(event.target.value) })}
                    aria-label="Annotation height"
                  />
                </label>
              </div>
            </div>
          ) : (
            <div className="empty-panel">
              <ImageIcon size={30} />
              <p>Select an annotation or choose a tool to add content to the page.</p>
            </div>
          )}

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
  onSelect,
  onAddMark,
  onUpdateMark,
}: {
  page: PageView;
  marks: Mark[];
  zoom: number;
  selectedMark: string | null;
  onSelect: (id: string | null) => void;
  onAddMark: (page: number, x: number, y: number) => void;
  onUpdateMark: (id: string, patch: Partial<Mark>) => void;
}) {
  const pageRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef<{ id: string; startX: number; startY: number; markX: number; markY: number } | null>(null);

  function pagePoint(event: PointerEvent<HTMLDivElement>) {
    const bounds = pageRef.current?.getBoundingClientRect();
    if (!bounds) return { x: 0, y: 0 };
    return {
      x: (event.clientX - bounds.left) / zoom,
      y: (event.clientY - bounds.top) / zoom,
    };
  }

  return (
    <div
      ref={pageRef}
      className="document-page"
      style={{ width: page.width * zoom, height: page.height * zoom }}
      onPointerDown={(event) => {
        if ((event.target as HTMLElement).closest(".mark")) return;
        const point = pagePoint(event);
        onSelect(null);
        onAddMark(page.pageNumber, point.x, point.y);
      }}
          onPointerMove={(event) => {
            const drag = dragRef.current;
            if (!drag) return;
            const activeMark = marks.find((mark) => mark.id === drag.id);
            if (!activeMark) return;
            const dx = (event.clientX - drag.startX) / zoom;
            const dy = (event.clientY - drag.startY) / zoom;
            onUpdateMark(drag.id, {
              x: clamp(drag.markX + dx, 0, Math.max(0, page.width - activeMark.width)),
              y: clamp(drag.markY + dy, 0, Math.max(0, page.height - activeMark.height)),
            });
          }}
      onPointerUp={() => {
        dragRef.current = null;
      }}
      onPointerLeave={() => {
        dragRef.current = null;
      }}
    >
      {page.imageUrl ? <img className="pdf-image" src={page.imageUrl} alt={`Page ${page.pageNumber}`} /> : <BlankPage />}
      {marks.map((mark) => (
        <button
          className={`mark mark-${mark.kind} ${selectedMark === mark.id ? "selected" : ""}`}
          key={mark.id}
          title={getMarkLabel(mark)}
          aria-label={getMarkLabel(mark)}
          style={{
            left: mark.x * zoom,
            top: mark.y * zoom,
            width: mark.width * zoom,
            height: mark.height * zoom,
            color: mark.color,
            fontSize: mark.size * zoom,
            backgroundColor: mark.kind === "highlight" ? hexToRgba(mark.color, 0.48) : undefined,
          }}
          onPointerDown={(event) => {
            event.stopPropagation();
            onSelect(mark.id);
            dragRef.current = {
              id: mark.id,
              startX: event.clientX,
              startY: event.clientY,
              markX: mark.x,
              markY: mark.y,
            };
          }}
        >
          {mark.kind === "highlight" ? "" : mark.text}
        </button>
      ))}
    </div>
  );
}

function BlankPage() {
  return (
    <div className="blank-page">
      <FilePlus2 size={44} />
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
        <strong>{message}</strong>
        <div className="progress-track" aria-label={progress === undefined ? message : `${Math.round(progress)}% complete`}>
          <span style={{ width: `${clamp(progress ?? 100, 0, 100)}%` }} />
        </div>
      </div>
    </div>
  );
}

function colorToRgb(hex: string) {
  const clean = hex.replace("#", "");
  const bigint = Number.parseInt(clean, 16);
  const red = ((bigint >> 16) & 255) / 255;
  const green = ((bigint >> 8) & 255) / 255;
  const blue = (bigint & 255) / 255;
  return rgb(red, green, blue);
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

function getMarkLabel(mark: Mark) {
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
