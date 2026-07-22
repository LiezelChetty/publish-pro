import type { DocumentBookmark, PageReference, TocLine } from "./types";

export function validateExportLinks(bookmarks: DocumentBookmark[], tocLines: TocLine[], pages: PageReference[]) {
  const pageIds = new Set(pages.map((page) => page.id));
  return {
    brokenBookmarkIds: bookmarks.filter((bookmark) => !pageIds.has(bookmark.pageId)).map((bookmark) => bookmark.id),
    brokenTocEntryIds: tocLines.filter((line) => !pageIds.has(line.pageId)).map((line) => line.id),
  };
}
