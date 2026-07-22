import { formatNumberingTemplate } from "./tokens";
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
  documentPages: number;
  numberedPages: number;
  sectionPages: number;
  included: boolean;
  includeInTotal: boolean;
};

export function resolveNumberingSections(pages: NumberingPage[], sections: NumberingSection[], customTemplate = "{page}") {
  const sortedSections = sortNumberingSections(pages, sections);
  const pageIndex = new Map(pages.map((page, index) => [page.id, index]));
  const sectionRanges = getNumberingSectionRanges(pages, sortedSections);
  const sectionTotals = new Map(sectionRanges.map(({ section, start, end }) => [section.id, section.includeNumbering ? Math.max(0, end - start + 1) : 0]));
  const numberedPages = sectionRanges.reduce((total, { section, start, end }) => {
    if (!section.includeNumbering || !section.includeInTotal) return total;
    return total + Math.max(0, end - start + 1);
  }, 0);
  const results = new Map<string, ResolvedPageNumber>();
  let previousValue = 0;
  for (const { section, start, end } of sectionRanges) {
    let value = section.restart ? section.startValue : previousValue + 1;
    for (let index = start; index <= Math.min(end, pages.length - 1); index += 1) {
      const page = pages[index];
      const sectionPages = sectionTotals.get(section.id) ?? 0;
      const display = section.usePageLabel && page.label ? page.label : section.includeNumbering ? formatNumberingTemplate({
        format: section.format,
        value,
        documentPages: pages.length,
        numberedPages,
        sectionPages,
        customTemplate,
        prefix: section.prefix,
        suffix: section.suffix,
      }) : "";
      results.set(page.id, {
        pageId: page.id,
        sectionId: section.id,
        sectionLabel: section.label,
        display,
        value,
        documentPages: pages.length,
        numberedPages,
        sectionPages,
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
        documentPages: pages.length,
        numberedPages: pages.length,
        sectionPages: pages.length,
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

function getNumberingSectionRanges(pages: NumberingPage[], sortedSections: NumberingSection[]) {
  const pageIndex = new Map(pages.map((page, index) => [page.id, index]));
  return sortedSections.flatMap((section, sectionIndex) => {
    const start = pageIndex.get(section.startPageId);
    if (start === undefined) return [];
    const explicitEnd = section.endPageId ? pageIndex.get(section.endPageId) : undefined;
    const nextStart = sortedSections[sectionIndex + 1] ? pageIndex.get(sortedSections[sectionIndex + 1].startPageId) : undefined;
    const end = explicitEnd ?? (nextStart === undefined ? pages.length - 1 : Math.max(start, nextStart - 1));
    return [{ section, start, end: Math.min(Math.max(start, end), pages.length - 1) }];
  });
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
