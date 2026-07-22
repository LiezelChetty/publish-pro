import { createDefaultPublishingSettings, mergeHeaderFooterZone } from "../publishing/defaults";
import { readDocxPackage } from "./docxPackage";
import { createBookmarksFromDocxHeadings } from "./docxBookmarks";
import { parseDocxContent, parseHeaderFooterText } from "./docxContent";
import { extractDocxImages, parseDocumentRelationships } from "./docxImages";
import { renderDocxToPdf } from "./docxLayout";
import { parseDocxNumbering, parseDocxStyles } from "./docxStyles";
import type { DocxHeading, DocxImportOptions, DocxImportProgress, DocxImportResult, DocxIntermediateDocument } from "./types";

export async function importDocxDocument(file: File, options: DocxImportOptions, onProgress?: (progress: DocxImportProgress) => void): Promise<DocxImportResult> {
  onProgress?.({ stage: "Reading document", progress: 5 });
  const originalBytes = new Uint8Array(await file.arrayBuffer());
  const projectName = file.name.replace(/\.docx$/i, "") || "Imported Word document";

  onProgress?.({ stage: "Opening DOCX package", progress: 12 });
  const pkg = readDocxPackage(originalBytes);
  const warnings = [...pkg.warnings];

  onProgress?.({ stage: "Parsing styles", progress: 22 });
  const styles = parseDocxStyles(pkg);
  const numbering = parseDocxNumbering(pkg);
  const relationships = parseDocumentRelationships(pkg);

  onProgress?.({ stage: "Extracting images", progress: 32 });
  const images = await extractDocxImages(pkg, relationships, warnings);

  onProgress?.({ stage: "Building document structure", progress: 45 });
  const content = parseDocxContent(pkg, styles, numbering, relationships, warnings, options.fallbackPageSize);
  const headerFooter: { headerText?: string; footerText?: string } = options.importHeadersFooters ? parseHeaderFooterText(pkg, relationships) : {};
  const intermediate: DocxIntermediateDocument = {
    title: projectName,
    blocks: content.blocks,
    pageSize: content.pageSize,
    images,
    headings: [],
    headerText: headerFooter.headerText,
    footerText: headerFooter.footerText,
    hyperlinks: content.hyperlinks,
    statistics: { ...content.statistics, imageCount: images.length },
    warnings,
  };

  if (images.some((image) => image.mimeType === "image/gif")) warnings.push({ code: "gif-fallback", message: "GIF images are preserved as Project Assets; first-phase PDF rendering does not animate or embed GIF frames." });
  if (images.some((image) => image.mimeType === "image/svg+xml")) warnings.push({ code: "svg-fallback", message: "SVG images are preserved as Project Assets; first-phase PDF rendering skips SVG placement when pdf-lib cannot embed them directly." });
  if (content.hyperlinks.length > 0 && options.importHyperlinks) warnings.push({ code: "hyperlink-appearance", message: "Hyperlink text is imported with visible styling. Real exported link annotations for DOCX-imported links are planned for a later phase." });

  onProgress?.({ stage: "Paginating Word content", progress: 62 });
  const rendered = await renderDocxToPdf(intermediate, options.fallbackFont);
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

  return {
    projectName,
    defaultPdfName: `${projectName}.pdf`,
    convertedPdfSource: {
      id: crypto.randomUUID(),
      name: `${projectName}.pdf`,
      bytes: rendered.bytes,
      mimeType: "application/pdf",
    },
    originalSource: options.preserveSource
      ? {
          id: crypto.randomUUID(),
          name: file.name,
          bytes: originalBytes,
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }
      : undefined,
    imageAssets: images,
    headings: normalizedHeadings,
    bookmarks,
    publishingSettingsPatch: publishingSettings,
    report: {
      pagesCreated: rendered.pageCount,
      headingsFound: normalizedHeadings.length,
      bookmarksCreated: bookmarks.length,
      imagesImported: images.length,
      tablesImported: content.statistics.tableCount,
      listsImported: content.statistics.listCount,
      headersFootersDetected: [headerFooter.headerText, headerFooter.footerText].filter(Boolean).length,
      footnotesDetected: content.statistics.footnoteCount,
      warnings,
      unsupportedContent: warnings.map((warning) => warning.message),
    },
  };
}

function normalizeHeadings(headings: DocxHeading[]) {
  return headings.filter((heading) => heading.title.trim()).map((heading) => ({ ...heading, level: Math.min(Math.max(heading.level, 1), 3) }));
}
