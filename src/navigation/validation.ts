import type { DocumentBookmark, PageReference, TocGeneratedPageMetadata, TocManualEntry } from "./types";

export type NavigationIssue = {
  id: string;
  severity: "warning" | "error";
  message: string;
  bookmarkId?: string;
  manualEntryId?: string;
};

export function validateNavigation({
  bookmarks,
  manualEntries,
  pages,
  generatedToc,
}: {
  bookmarks: DocumentBookmark[];
  manualEntries: TocManualEntry[];
  pages: PageReference[];
  generatedToc?: TocGeneratedPageMetadata;
}) {
  const issues: NavigationIssue[] = [];
  const pageIds = new Set(pages.map((page) => page.id));
  const bookmarkIds = new Set<string>();
  for (const bookmark of bookmarks) {
    if (bookmarkIds.has(bookmark.id)) issues.push({ id: `duplicate-${bookmark.id}`, severity: "error", message: `Duplicate bookmark ID: ${bookmark.id}`, bookmarkId: bookmark.id });
    bookmarkIds.add(bookmark.id);
    if (!pageIds.has(bookmark.pageId)) issues.push({ id: `broken-${bookmark.id}`, severity: "error", message: `"${bookmark.title}" points to a missing page.`, bookmarkId: bookmark.id });
    if (bookmark.parentId && !bookmarks.some((item) => item.id === bookmark.parentId)) issues.push({ id: `missing-parent-${bookmark.id}`, severity: "warning", message: `"${bookmark.title}" has a missing parent.`, bookmarkId: bookmark.id });
    if (hasCycle(bookmarks, bookmark.id)) issues.push({ id: `cycle-${bookmark.id}`, severity: "error", message: `"${bookmark.title}" has a circular parent relationship.`, bookmarkId: bookmark.id });
  }
  for (const entry of manualEntries) {
    if (entry.enabled && !pageIds.has(entry.pageId)) issues.push({ id: `manual-broken-${entry.id}`, severity: "error", message: `Manual TOC entry "${entry.title}" points to a missing page.`, manualEntryId: entry.id });
  }
  if (generatedToc) {
    for (const pageId of generatedToc.pageIds) {
      if (!pageIds.has(pageId)) issues.push({ id: `toc-page-${pageId}`, severity: "warning", message: "A generated TOC page is missing from the document." });
    }
  }
  return issues;
}

function hasCycle(bookmarks: DocumentBookmark[], bookmarkId: string) {
  const visited = new Set<string>();
  let current = bookmarks.find((bookmark) => bookmark.id === bookmarkId);
  while (current?.parentId) {
    if (visited.has(current.parentId)) return true;
    visited.add(current.parentId);
    current = bookmarks.find((bookmark) => bookmark.id === current?.parentId);
  }
  return false;
}
