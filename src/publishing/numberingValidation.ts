import type { NumberingSection } from "./numberingSections";
import type { NumberingPage } from "./numberingResolver";

export type NumberingIssue = {
  id: string;
  severity: "warning" | "error";
  message: string;
  sectionId?: string;
};

export function validateNumberingSections(pages: NumberingPage[], sections: NumberingSection[]) {
  const issues: NumberingIssue[] = [];
  const pageIndex = new Map(pages.map((page, index) => [page.id, index]));
  const sectionIds = new Set<string>();
  for (const section of sections) {
    if (sectionIds.has(section.id)) issues.push({ id: `duplicate-${section.id}`, severity: "error", message: `Duplicate numbering section ID: ${section.id}`, sectionId: section.id });
    sectionIds.add(section.id);
    const start = pageIndex.get(section.startPageId);
    const end = section.endPageId ? pageIndex.get(section.endPageId) : undefined;
    if (start === undefined) issues.push({ id: `missing-start-${section.id}`, severity: "error", message: `${section.label || "Section"} is missing its start page.`, sectionId: section.id });
    if (section.endPageId && end === undefined) issues.push({ id: `missing-end-${section.id}`, severity: "warning", message: `${section.label || "Section"} is missing its end page.`, sectionId: section.id });
    if (start !== undefined && end !== undefined && end < start) issues.push({ id: `end-before-start-${section.id}`, severity: "error", message: `${section.label || "Section"} ends before it starts.`, sectionId: section.id });
    if (!Number.isFinite(section.startValue) || section.startValue < 1) issues.push({ id: `bad-start-${section.id}`, severity: "error", message: `${section.label || "Section"} has an invalid starting value.`, sectionId: section.id });
  }
  const ranges: Array<{ section: NumberingSection; start: number; end?: number }> = [];
  for (const section of sections) {
    const start = pageIndex.get(section.startPageId);
    if (start === undefined) continue;
    ranges.push({ section, start, end: section.endPageId ? pageIndex.get(section.endPageId) : undefined });
  }
  ranges.sort((a, b) => a.start - b.start);
  for (let index = 1; index < ranges.length; index += 1) {
    const previous = ranges[index - 1];
    const current = ranges[index];
    if (previous.end !== undefined && previous.end >= current.start) {
      issues.push({ id: `overlap-${previous.section.id}-${current.section.id}`, severity: "warning", message: `${previous.section.label} overlaps ${current.section.label}.`, sectionId: current.section.id });
    }
  }
  return issues;
}
