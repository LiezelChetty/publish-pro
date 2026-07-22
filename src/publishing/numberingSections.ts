import type { PublishingNumberFormat } from "./types";

export type NumberingSection = {
  id: string;
  label: string;
  startPageId: string;
  endPageId?: string;
  format: PublishingNumberFormat;
  startValue: number;
  prefix: string;
  suffix: string;
  includeNumbering: boolean;
  includeInTotal: boolean;
  restart: boolean;
  usePageLabel: boolean;
};

export function createNumberingSection(pageId: string, label = "Section"): NumberingSection {
  return {
    id: crypto.randomUUID(),
    label,
    startPageId: pageId,
    format: "decimal",
    startValue: 1,
    prefix: "",
    suffix: "",
    includeNumbering: true,
    includeInTotal: true,
    restart: true,
    usePageLabel: false,
  };
}

export function duplicateNumberingSection(section: NumberingSection) {
  return {
    ...section,
    id: crypto.randomUUID(),
    label: `${section.label} Copy`,
  };
}
