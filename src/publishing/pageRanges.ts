import type { PublishingPageLike, PublishingTarget } from "./types";

export function resolvePublishingPageIds(target: PublishingTarget, pages: PublishingPageLike[], currentPage: number, selectedPageIds: string[]) {
  const selected = new Set(selectedPageIds);
  let ids: string[];

  if (target.mode === "current") {
    ids = pages.filter((page) => page.pageNumber === currentPage).map((page) => page.id);
  } else if (target.mode === "selected") {
    ids = pages.filter((page) => selected.has(page.id)).map((page) => page.id);
  } else if (target.mode === "exceptSelected") {
    ids = pages.filter((page) => !selected.has(page.id)).map((page) => page.id);
  } else if (target.mode === "custom") {
    ids = parsePageRange(target.range, pages.length).map((pageNumber) => pages[pageNumber - 1]?.id).filter((id): id is string => Boolean(id));
  } else if (target.mode === "odd") {
    ids = pages.filter((page) => page.pageNumber % 2 === 1).map((page) => page.id);
  } else if (target.mode === "even") {
    ids = pages.filter((page) => page.pageNumber % 2 === 0).map((page) => page.id);
  } else if (target.mode === "exceptFirst") {
    ids = pages.filter((page) => page.pageNumber !== 1).map((page) => page.id);
  } else {
    ids = pages.map((page) => page.id);
  }

  if (target.excludeFirst) ids = ids.filter((id) => pages.find((page) => page.id === id)?.pageNumber !== 1);
  return new Set(ids);
}

export function getPageRangeError(range: string, pageCount: number) {
  try {
    parsePageRange(range, pageCount);
    return "";
  } catch (error) {
    return error instanceof Error ? error.message : "Invalid page range.";
  }
}

export function parsePageRange(range: string, pageCount: number) {
  const trimmed = range.trim();
  if (!trimmed) throw new Error("Enter a page range such as 1-5 or 1, 3, 7-9.");
  const result = new Set<number>();
  for (const part of trimmed.split(",")) {
    const segment = part.trim();
    if (!segment) continue;
    const [startText, endText] = segment.split("-").map((item) => item.trim());
    const start = Number(startText);
    const end = endText === undefined ? start : Number(endText);
    if (!Number.isInteger(start) || !Number.isInteger(end)) throw new Error(`Invalid page range segment: ${segment}.`);
    if (start < 1 || end < 1 || start > pageCount || end > pageCount) throw new Error(`Page range ${segment} is outside this ${pageCount}-page document.`);
    if (start > end) throw new Error(`Page range ${segment} starts after it ends.`);
    for (let page = start; page <= end; page += 1) result.add(page);
  }
  return [...result].sort((left, right) => left - right);
}
