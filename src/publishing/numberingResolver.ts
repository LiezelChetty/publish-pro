import { formatPageNumber } from "./tokens";
import type { PublishingNumberFormat } from "./types";
import type { NumberingSection } from "./numberingSections";

export type NumberingPage = {
  id: string;
  pageNumber: number;
  label?: string;
  generatedToc?: boolean;
};

export type ResolvedPageNumber = {
  pageId: string;
  sectionId?: string;
  sectionLabel?: string;
  display: string;
  value?: number;
  included: boolean;
  includeInTotal: boolean;
};

export function resolveNumberingSections(pages: NumberingPage[], sections: NumberingSection[], customTemplate = "{page}") {
  const sortedSections = sortNumberingSections(pages, sections);
  const pageIndex = new Map(pages.map((page, index) => [page.id, index]));
  const results = new Map<string, ResolvedPageNumber>();
  let previousValue = 0;
  for (const [sectionIndex, section] of sortedSections.entries()) {
    const start = pageIndex.get(section.startPageId);
    if (start === undefined) continue;
    const explicitEnd = section.endPageId ? pageIndex.get(section.endPageId) : undefined;
    const nextStart = sortedSections[sectionIndex + 1] ? pageIndex.get(sortedSections[sectionIndex + 1].startPageId) : undefined;
    const end = explicitEnd ?? (nextStart === undefined ? pages.length - 1 : Math.max(start, nextStart - 1));
    let value = section.restart ? section.startValue : previousValue + 1;
    for (let index = start; index <= Math.min(end, pages.length - 1); index += 1) {
      const page = pages[index];
      const display = section.usePageLabel && page.label ? page.label : section.includeNumbering ? `${section.prefix}${formatPageNumber(section.format, value, 0, customTemplate)}${section.suffix}` : "";
      results.set(page.id, {
        pageId: page.id,
        sectionId: section.id,
        sectionLabel: section.label,
        display,
        value,
        included: section.includeNumbering,
        includeInTotal: section.includeInTotal,
      });
      if (section.includeNumbering) {
        previousValue = value;
        value += 1;
      }
    }
  }
  for (const page of pages) {
    if (!results.has(page.id)) {
      results.set(page.id, {
        pageId: page.id,
        display: String(page.pageNumber),
        value: page.pageNumber,
        included: true,
        includeInTotal: true,
      });
    }
  }
  return results;
}

export function getResolvedPageText(page: NumberingPage, resolved: Map<string, ResolvedPageNumber>) {
  return resolved.get(page.id)?.display || "";
}

export function sortNumberingSections(pages: NumberingPage[], sections: NumberingSection[]) {
  const pageIndex = new Map(pages.map((page, index) => [page.id, index]));
  return [...sections].sort((a, b) => (pageIndex.get(a.startPageId) ?? Number.MAX_SAFE_INTEGER) - (pageIndex.get(b.startPageId) ?? Number.MAX_SAFE_INTEGER));
}

export function getNumberingSummary(pages: NumberingPage[], sections: NumberingSection[], customTemplate = "{page}") {
  const resolved = resolveNumberingSections(pages, sections, customTemplate);
  return sortNumberingSections(pages, sections).map((section) => {
    const sectionPages = pages.filter((page) => resolved.get(page.id)?.sectionId === section.id);
    const first = sectionPages[0] ? resolved.get(sectionPages[0].id)?.display : "";
    const last = sectionPages[sectionPages.length - 1] ? resolved.get(sectionPages[sectionPages.length - 1].id)?.display : "";
    return {
      section,
      text: section.includeNumbering ? `${section.label || "Section"} - ${first}${last && last !== first ? `-${last}` : ""}` : `${section.label || "Section"} - unnumbered`,
    };
  });
}

export function defaultLegacyNumberingSection(pageId: string, format: PublishingNumberFormat, startValue: number, prefix: string, suffix: string): NumberingSection {
  return {
    id: crypto.randomUUID(),
    label: "Document",
    startPageId: pageId,
    format,
    startValue,
    prefix,
    suffix,
    includeNumbering: true,
    includeInTotal: true,
    restart: true,
    usePageLabel: false,
  };
}
