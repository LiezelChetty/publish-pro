import { createDefaultPublishingSettings, mergeHeaderFooterZone } from "../publishing/defaults";
import { readDocxPackage } from "./docxPackage";
import { createBookmarksFromDocxHeadings } from "./docxBookmarks";
import { parseDocxContent, parseHeaderFooterText } from "./docxContent";
import { extractDocxImages, parseDocumentRelationships } from "./docxImages";
import { renderDocxToPdf } from "./docxLayout";
import { collectStyleWarnings, parseDocxNumbering, parseDocxStyles } from "./docxStyles";
import type { DocxHeading, DocxImportOptions, DocxImportProgress, DocxImportResult, DocxIntermediateDocument } from "./types";

export async function importDocxDocument(file: File, options: DocxImportOptions, onProgress?: (progress: DocxImportProgress) => void): Promise<DocxImportResult> {
  const startedAt = performance.now();
  const importId = crypto.randomUUID();
  onProgress?.({ stage: "Reading document", progress: 5 });
  const originalBytes = new Uint8Array(await file.arrayBuffer());
  const projectName = file.name.replace(/\.docx$/i, "") || "Imported Word document";

  onProgress?.({ stage: "Opening DOCX package", progress: 12 });
  const pkg = readDocxPackage(originalBytes);
  const warnings = [...pkg.warnings];

  onProgress?.({ stage: "Parsing styles", progress: 22 });
  const parseStartedAt = performance.now();
  const styles = parseDocxStyles(pkg);
  collectStyleWarnings(styles, warnings);
  const numbering = parseDocxNumbering(pkg);
  const relationships = parseDocumentRelationships(pkg);

  onProgress?.({ stage: "Extracting images", progress: 32 });
  const images = await extractDocxImages(pkg, relationships, warnings);

  onProgress?.({ stage: "Building document structure", progress: 45 });
  const content = parseDocxContent(pkg, styles, numbering, relationships, warnings, options);
  const headerFooter: { headerText?: string; footerText?: string } = options.importHeadersFooters ? parseHeaderFooterText(pkg, relationships) : {};
  const parseTimeMs = performance.now() - parseStartedAt;
  const intermediate: DocxIntermediateDocument = {
    title: projectName,
    blocks: content.blocks,
    pageSize: content.pageSize,
    images,
    headings: [],
    headerText: headerFooter.headerText,
    footerText: headerFooter.footerText,
    hyperlinks: content.hyperlinks,
    notes: content.notes,
    wordBookmarks: content.wordBookmarks,
    internalLinks: content.internalLinks,
    wordComments: content.wordComments,
    trackedChanges: content.trackedChanges,
    sectionNumbering: content.sectionNumbering,
    statistics: { ...content.statistics, imageCount: images.length },
    warnings,
  };

  if (images.some((image) => image.mimeType === "image/gif")) warnings.push({ code: "gif-fallback", message: "GIF images are preserved as Project Assets; first-phase PDF rendering does not animate or embed GIF frames." });
  if (images.some((image) => image.mimeType === "image/svg+xml")) warnings.push({ code: "svg-fallback", message: "SVG images are preserved as Project Assets; first-phase PDF rendering skips SVG placement when pdf-lib cannot embed them directly." });
  if (content.internalLinks.length > 0 && content.wordBookmarks.length === 0) warnings.push({ code: "internal-hyperlink-fallback", category: "hyperlinks", message: "Some internal Word hyperlinks were detected without matching bookmark markers. They will target the top of the document when no heading or bookmark can be resolved." });

  onProgress?.({ stage: "Paginating Word content", progress: 62 });
  const renderStartedAt = performance.now();
  const rendered = await renderDocxToPdf(intermediate, options.fallbackFont, options.fidelityMode);
  const renderTimeMs = performance.now() - renderStartedAt;
  const pageIdsByNumber = new Map<number, string>();
  for (let pageNumber = 1; pageNumber <= rendered.pageCount; pageNumber += 1) {
    pageIdsByNumber.set(pageNumber, `pending-page-${pageNumber}`);
  }
  const normalizedHeadings = normalizeHeadings(rendered.headings);

  onProgress?.({ stage: "Creating bookmarks", progress: 82 });
  const bookmarks = options.createBookmarks ? createBookmarksFromDocxHeadings(normalizedHeadings, pageIdsByNumber) : [];

  onProgress?.({ stage: "Finalising import", progress: 95 });
  const publishingSettings = createDefaultPublishingSettings();
  if (options.importHeadersFooters && (headerFooter.headerText || headerFooter.footerText)) {
    publishingSettings.headerFooter.enabled = true;
    publishingSettings.headerFooter.header.left = mergeHeaderFooterZone(undefined, headerFooter.headerText ?? "");
    publishingSettings.headerFooter.footer.center = mergeHeaderFooterZone(undefined, headerFooter.footerText ?? "");
  }

  const convertedSourceId = crypto.randomUUID();
  const originalSourceId = options.preserveSource ? crypto.randomUUID() : undefined;
  const importedAt = new Date().toISOString();
  const sourceHash = await hashBytes(originalBytes);
  const revision = {
    revisionId: crypto.randomUUID(),
    sourceHash,
    importedAt,
    importerVersion: "docx-import-v2",
    fidelityMode: options.fidelityMode,
    trackedChangesMode: options.trackedChangesMode,
    pageCount: rendered.pageCount,
    warningCount: warnings.length,
  };
  const fontSubstitutions = warnings.filter((warning) => warning.code === "font-substitution").map((warning) => warning.message);
  const layoutSimplifications = warnings.filter((warning) => warning.code.includes("fallback") || warning.code.includes("simplified") || warning.code === "table-scaled").map((warning) => warning.message);

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
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }
      : undefined,
    imageAssets: images,
    headings: normalizedHeadings,
    bookmarks,
    links: rendered.links,
    sourceMappings: rendered.sourceMappings,
    wordComments: content.wordComments,
    sectionNumbering: content.sectionNumbering,
    importMetadata: {
      importId,
      sourceDocumentId: convertedSourceId,
      originalSourceDocumentId: originalSourceId,
      sourceName: file.name,
      importedAt,
      fidelityMode: options.fidelityMode,
      options,
      pageCount: rendered.pageCount,
      warningCount: warnings.length,
      revisionHistory: [revision],
    },
    publishingSettingsPatch: publishingSettings,
    report: {
      importId,
      fidelityMode: options.fidelityMode,
      pagesCreated: rendered.pageCount,
      sectionsDetected: content.statistics.sectionCount,
      headingsFound: normalizedHeadings.length,
      bookmarksCreated: bookmarks.length,
      imagesImported: images.length,
      tablesImported: content.statistics.tableCount,
      listsImported: content.statistics.listCount,
      hyperlinksImported: rendered.links.length,
      headersFootersDetected: [headerFooter.headerText, headerFooter.footerText].filter(Boolean).length,
      footnotesDetected: content.statistics.footnoteCount,
      endnotesDetected: content.statistics.endnoteCount ?? 0,
      notesPlacedExactly: content.statistics.notesPlacedExactly ?? 0,
      notesUsingFallback: content.statistics.notesUsingFallback ?? 0,
      brokenNoteReferences: content.statistics.brokenNoteReferences ?? 0,
      commentsDetected: content.statistics.commentsDetected,
      commentsImported: content.statistics.commentsImported ?? 0,
      approximateComments: content.statistics.approximateComments ?? 0,
      trackedChangesDetected: content.statistics.trackedChangesDetected,
      trackedChangesMode: options.trackedChangesMode,
      trackedChangeSummary: content.trackedChanges,
      internalLinksImported: rendered.links.filter((link) => link.kind === "internal").length,
      internalLinkFallbacks: rendered.links.filter((link) => link.kind === "internal" && link.usedFallbackDestination).length,
      sectionNumberingDetected: content.sectionNumbering.length,
      revision,
      fontSubstitutions,
      layoutSimplifications,
      sourceMappings: rendered.sourceMappings.length,
      parseTimeMs,
      renderTimeMs,
      totalTimeMs: performance.now() - startedAt,
      warnings,
      unsupportedContent: warnings.map((warning) => warning.message),
    },
  };
}

function normalizeHeadings(headings: DocxHeading[]) {
  return headings.filter((heading) => heading.title.trim()).map((heading) => ({ ...heading, level: Math.min(Math.max(heading.level, 1), 3) }));
}

async function hashBytes(bytes: Uint8Array) {
  const digest = await crypto.subtle.digest("SHA-256", bytes.slice().buffer);
  return Array.from(new Uint8Array(digest)).map((byte) => byte.toString(16).padStart(2, "0")).join("");
}
