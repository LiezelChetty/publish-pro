import { getBookmarkLevel } from "./bookmarks";
import type { DocumentBookmark, PageReference, TocEntry, TocLayoutResult, TocManualEntry, TocSettings } from "./types";

export function getTocPageSize(size: TocSettings["pageSize"]) {
  return size === "letter" ? { width: 612, height: 792 } : { width: 595.28, height: 841.89 };
}

export function buildTocEntries(bookmarks: DocumentBookmark[], manualEntries: TocManualEntry[], pages: PageReference[], settings: TocSettings): TocEntry[] {
  const pageIds = new Set(pages.map((page) => page.id));
  const entries: TocEntry[] = [];
  if (settings.source === "bookmarks") {
    entries.push(
      ...bookmarks
        .filter((bookmark) => pageIds.has(bookmark.pageId))
        .sort((a, b) => getBookmarkSortKey(bookmarks, a).localeCompare(getBookmarkSortKey(bookmarks, b)))
        .map((bookmark) => ({
          id: bookmark.id,
          title: bookmark.title,
          pageId: bookmark.pageId,
          level: getBookmarkLevel(bookmarks, bookmark),
          source: "bookmark" as const,
        }))
        .filter((entry) => entry.level <= settings.maxDepth)
    );
  }
  if (settings.source === "pageLabels") {
    entries.push(
      ...pages.map((page) => ({
        id: page.id,
        title: page.label || `Page ${page.pageNumber}`,
        pageId: page.id,
        level: 1,
        source: "pageLabel" as const,
      }))
    );
  }
  entries.push(
    ...manualEntries
      .filter((entry) => entry.enabled && pageIds.has(entry.pageId))
      .sort((a, b) => a.order - b.order)
      .map((entry) => ({
        id: entry.id,
        title: entry.title,
        pageId: entry.pageId,
        level: Math.max(1, entry.level),
        source: "manual" as const,
      }))
  );
  return entries;
}

export function layoutToc(entries: TocEntry[], pages: PageReference[], settings: TocSettings): TocLayoutResult {
  const pageSize = getTocPageSize(settings.pageSize);
  const lineHeight = settings.fontSize * settings.lineSpacing;
  const usableWidth = pageSize.width - settings.marginLeft - settings.marginRight;
  const usableHeight = pageSize.height - settings.marginTop - settings.marginBottom - settings.fontSize * 2.6;
  const pageById = new Map(pages.map((page) => [page.id, page]));
  const result: TocLayoutResult = { pages: [{ id: crypto.randomUUID(), lines: [] }], width: pageSize.width, height: pageSize.height };
  let pageIndex = 0;
  let cursorY = pageSize.height - settings.marginTop - settings.fontSize * 2.6;
  for (const entry of entries) {
    const page = pageById.get(entry.pageId);
    const pageText = page ? getTocPageText(page, settings) : "Missing";
    const indent = Math.max(0, entry.level - 1) * settings.indent;
    const maxTitleChars = Math.max(16, Math.floor((usableWidth - indent - 48) / (settings.fontSize * 0.52)));
    const wrapped = wrapTitle(entry.title, maxTitleChars);
    const requiredHeight = Math.max(1, wrapped.length) * lineHeight;
    if (cursorY - requiredHeight < settings.marginBottom && result.pages[pageIndex].lines.length > 0) {
      result.pages.push({ id: crypto.randomUUID(), lines: [] });
      pageIndex += 1;
      cursorY = pageSize.height - settings.marginTop - settings.fontSize * 2.6;
    }
    wrapped.forEach((text, wrapIndex) => {
      result.pages[pageIndex].lines.push({
        ...entry,
        text,
        pageText: wrapIndex === wrapped.length - 1 ? pageText : "",
        x: settings.marginLeft + indent,
        y: cursorY,
        pageIndex,
        width: usableWidth - indent,
        height: lineHeight,
      });
      cursorY -= lineHeight;
    });
  }
  if (result.pages.length === 0) result.pages.push({ id: crypto.randomUUID(), lines: [] });
  return result;
}

export function getTocPageText(page: PageReference, settings: TocSettings) {
  if (settings.includePageLabels && page.label) return page.label;
  if (settings.includePageNumbers) return String(page.pageNumber);
  return "";
}

function getBookmarkSortKey(bookmarks: DocumentBookmark[], bookmark: DocumentBookmark): string {
  const path: string[] = [`${bookmark.order}`.padStart(5, "0")];
  let parent = bookmarks.find((item) => item.id === bookmark.parentId);
  while (parent) {
    path.unshift(`${parent.order}`.padStart(5, "0"));
    parent = bookmarks.find((item) => item.id === parent?.parentId);
  }
  return path.join(".");
}

function wrapTitle(title: string, maxChars: number) {
  const words = title.split(/\s+/).filter(Boolean);
  if (words.length === 0) return ["Untitled"];
  const lines: string[] = [];
  let current = "";
  for (const word of words) {
    if (!current) {
      current = word;
    } else if (`${current} ${word}`.length <= maxChars) {
      current = `${current} ${word}`;
    } else {
      lines.push(current);
      current = word;
    }
  }
  if (current) lines.push(current);
  return lines;
}
