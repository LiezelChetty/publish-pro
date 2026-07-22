import { isDocxFile, rejectUnsupportedWordFile } from "./docxPackage";
import type { DocxImportOptions } from "./types";

export const defaultDocxImportOptions: DocxImportOptions = {
  fidelityMode: "balanced",
  createBookmarks: true,
  preserveSource: true,
  importHeadersFooters: true,
  importHyperlinks: true,
  rebuildTocFromHeadings: true,
  fallbackPageSize: "letter",
  fallbackFont: "Helvetica",
  trackedChangesMode: "accepted",
};

export function validateDocxImportFile(file: File) {
  rejectUnsupportedWordFile(file);
  if (!isDocxFile(file)) throw new Error("Choose a .docx Word document.");
  if (file.size <= 0) throw new Error("The selected DOCX file is empty.");
  if (file.size > 40 * 1024 * 1024) {
    throw new Error("This DOCX is unusually large for the first-phase importer. Try a smaller document or split it before importing.");
  }
}
