import { readPptxPackage } from "./pptxPackage";
import { parsePptxPresentation } from "./pptxReader";
import { renderPptxToPdf } from "./pptxLayout";
import type { PptxImportOptions, PptxImportProgress, PptxImportResult } from "./pptxTypes";

export async function importPptxPresentation(file: File, options: PptxImportOptions, onProgress?: (progress: PptxImportProgress) => void): Promise<PptxImportResult> {
  const startedAt = performance.now();
  const importId = crypto.randomUUID();
  onProgress?.({ stage: "Reading presentation", progress: 6 });
  const originalBytes = new Uint8Array(await file.arrayBuffer());
  const projectName = file.name.replace(/\.pptx$/i, "") || "Imported PowerPoint presentation";

  onProgress?.({ stage: "Opening PPTX package", progress: 14 });
  const parseStartedAt = performance.now();
  const pkg = readPptxPackage(originalBytes);

  onProgress?.({ stage: "Resolving themes and layouts", progress: 28 });
  const presentation = parsePptxPresentation(pkg, options);
  const parseTimeMs = performance.now() - parseStartedAt;
  if (options.fidelityMode === "high") {
    presentation.warnings.push({ code: "pptx-high-fidelity-limited", category: "general", message: "High Fidelity improves source metadata preservation, but exact PowerPoint rendering still differs in this first phase." });
  }

  onProgress?.({ stage: "Rendering slides", progress: 56 });
  const renderStartedAt = performance.now();
  const rendered = await renderPptxToPdf(presentation);
  const renderTimeMs = performance.now() - renderStartedAt;

  onProgress?.({ stage: "Creating bookmarks", progress: 82 });
  const pageIdsByNumber = new Map<number, string>();
  for (let pageNumber = 1; pageNumber <= rendered.pageCount; pageNumber += 1) pageIdsByNumber.set(pageNumber, `pending-page-${pageNumber}`);
  const bookmarks = options.createBookmarks ? presentation.slides.map((slide) => ({
    id: crypto.randomUUID(),
    title: slide.title,
    pageId: pageIdsByNumber.get(slide.slideIndex) ?? `pending-page-${slide.slideIndex}`,
    order: slide.slideIndex,
    expanded: true,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  })) : [];

  onProgress?.({ stage: "Finalising project", progress: 95 });
  const convertedSourceId = crypto.randomUUID();
  const originalSourceId = options.preserveSource ? crypto.randomUUID() : undefined;
  const importedAt = new Date().toISOString();
  const revision = {
    revisionId: crypto.randomUUID(),
    sourceHash: await hashBytes(originalBytes),
    importedAt,
    importerVersion: "pptx-import-v1",
    fidelityMode: options.fidelityMode,
    slideCount: rendered.pageCount,
    warningCount: presentation.warnings.length,
  };

  return {
    projectName,
    defaultPdfName: `${projectName}.pdf`,
    convertedPdfSource: {
      id: convertedSourceId,
      name: `${projectName}.pdf`,
      bytes: rendered.bytes,
      mimeType: "application/pdf",
    },
    originalSource: options.preserveSource
      ? {
          id: originalSourceId!,
          name: file.name,
          bytes: originalBytes,
          mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        }
      : undefined,
    imageAssets: presentation.imageAssets,
    bookmarks,
    sourceMappings: rendered.sourceMappings,
    importMetadata: {
      kind: "pptx",
      importId,
      sourceDocumentId: convertedSourceId,
      originalSourceDocumentId: originalSourceId,
      sourceName: file.name,
      importedAt,
      fidelityMode: options.fidelityMode,
      options,
      pageCount: rendered.pageCount,
      warningCount: presentation.warnings.length,
      revisionHistory: [revision],
      slideTitles: presentation.slides.map((slide) => ({ pageNumber: slide.slideIndex, title: slide.title, hidden: slide.hidden })),
      speakerNotes: presentation.slides.filter((slide) => slide.notes.trim()).map((slide) => ({ pageNumber: slide.slideIndex, slideId: slide.id, title: slide.title, notes: slide.notes })),
      sections: presentation.sections,
    },
    report: {
      importId,
      fidelityMode: options.fidelityMode,
      slidesImported: rendered.pageCount,
      hiddenSlides: presentation.statistics.hiddenSlides,
      slideSections: presentation.sections.length,
      textBoxes: presentation.statistics.textBlocks,
      images: presentation.statistics.images,
      shapes: presentation.statistics.shapes,
      tables: presentation.statistics.tables,
      charts: presentation.statistics.charts,
      smartArtFallbacks: presentation.warnings.filter((warning) => warning.code.includes("smartart")).length,
      speakerNotes: presentation.statistics.speakerNotes,
      hyperlinks: presentation.statistics.hyperlinks,
      fontSubstitutions: ["PowerPoint theme fonts use Helvetica fallback in first-phase import."],
      rasterisedElements: presentation.statistics.rasterisedElements,
      unsupportedElements: presentation.statistics.unsupportedElements,
      parseTimeMs,
      renderTimeMs,
      totalTimeMs: performance.now() - startedAt,
      revision,
      warnings: presentation.warnings,
    },
  };
}

async function hashBytes(bytes: Uint8Array) {
  const digest = await crypto.subtle.digest("SHA-256", bytes.slice().buffer);
  return Array.from(new Uint8Array(digest)).map((byte) => byte.toString(16).padStart(2, "0")).join("");
}
