import { isPptxFile, rejectUnsupportedPowerPointFile } from "./pptxPackage";
import type { PptxImportOptions } from "./pptxTypes";

export const defaultPptxImportOptions: PptxImportOptions = {
  fidelityMode: "balanced",
  preserveSource: true,
  createBookmarks: true,
  includeHiddenSlides: true,
};

export function validatePptxImportFile(file: File) {
  rejectUnsupportedPowerPointFile(file);
  if (!isPptxFile(file)) throw new Error("Choose a .pptx PowerPoint presentation.");
  if (file.size <= 0) throw new Error("The selected PPTX file is empty.");
  if (file.size > 80 * 1024 * 1024) throw new Error("This PPTX is unusually large for the first-phase importer. Try a smaller deck or split it before importing.");
}
