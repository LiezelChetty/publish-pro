import type { DocumentBookmark, PageReference, TocSettings } from "./types";

export function createDefaultTocSettings(): TocSettings {
  return {
    title: "Table of Contents",
    source: "bookmarks",
    includePageLabels: true,
    includePageNumbers: true,
    dotLeaders: true,
    maxDepth: 4,
    indent: 18,
    lineSpacing: 1.35,
    fontFamily: "Helvetica",
    fontSize: 11,
    color: "#111827",
    boldLevels: [1],
    marginTop: 72,
    marginRight: 54,
    marginBottom: 54,
    marginLeft: 54,
    pageSize: "a4",
    insertPosition: "beforeFirst",
    includeTocInNumbering: true,
    startMainNumberingAfterToc: false,
  };
}

export function createBookmark(page: PageReference, title?: string, parentId?: string, order = 0): DocumentBookmark {
  const now = new Date().toISOString();
  return {
    id: crypto.randomUUID(),
    title: title?.trim() || page.label || `Page ${page.pageNumber}`,
    pageId: page.id,
    parentId,
    order,
    expanded: true,
    createdAt: now,
    updatedAt: now,
  };
}

export function normalizeBookmarks(bookmarks: DocumentBookmark[]) {
  const byParent = new Map<string, DocumentBookmark[]>();
  for (const bookmark of bookmarks) {
    const parentKey = bookmark.parentId ?? "";
    byParent.set(parentKey, [...(byParent.get(parentKey) ?? []), bookmark]);
  }
  return bookmarks.map((bookmark) => {
    const siblings = (byParent.get(bookmark.parentId ?? "") ?? []).sort((a, b) => a.order - b.order || a.createdAt.localeCompare(b.createdAt));
    return { ...bookmark, order: siblings.findIndex((item) => item.id === bookmark.id) };
  });
}

export function duplicateBookmarkTree(bookmarks: DocumentBookmark[], bookmarkId: string) {
  const source = bookmarks.find((bookmark) => bookmark.id === bookmarkId);
  if (!source) return bookmarks;
  const descendants = getBookmarkDescendants(bookmarks, bookmarkId);
  const idMap = new Map<string, string>([[source.id, crypto.randomUUID()]]);
  descendants.forEach((bookmark) => idMap.set(bookmark.id, crypto.randomUUID()));
  const now = new Date().toISOString();
  const copies = [source, ...descendants].map((bookmark) => ({
    ...bookmark,
    id: idMap.get(bookmark.id) ?? crypto.randomUUID(),
    parentId: bookmark.id === source.id ? source.parentId : idMap.get(bookmark.parentId ?? ""),
    title: bookmark.id === source.id ? `${bookmark.title} Copy` : bookmark.title,
    order: bookmark.id === source.id ? source.order + 1 : bookmark.order,
    createdAt: now,
    updatedAt: now,
  }));
  return normalizeBookmarks([...bookmarks, ...copies]);
}

export function getBookmarkDescendants(bookmarks: DocumentBookmark[], bookmarkId: string): DocumentBookmark[] {
  const children = bookmarks.filter((bookmark) => bookmark.parentId === bookmarkId).sort((a, b) => a.order - b.order);
  return children.flatMap((child) => [child, ...getBookmarkDescendants(bookmarks, child.id)]);
}

export function wouldCreateBookmarkCycle(bookmarks: DocumentBookmark[], bookmarkId: string, nextParentId?: string) {
  if (!nextParentId) return false;
  if (bookmarkId === nextParentId) return true;
  return getBookmarkDescendants(bookmarks, bookmarkId).some((bookmark) => bookmark.id === nextParentId);
}

export function getBookmarkLevel(bookmarks: DocumentBookmark[], bookmark: DocumentBookmark) {
  let level = 1;
  let parentId = bookmark.parentId;
  const visited = new Set<string>([bookmark.id]);
  while (parentId) {
    if (visited.has(parentId)) return level;
    visited.add(parentId);
    const parent = bookmarks.find((item) => item.id === parentId);
    if (!parent) return level;
    level += 1;
    parentId = parent.parentId;
  }
  return level;
}

export function getVisibleBookmarkTree(bookmarks: DocumentBookmark[]) {
  const sorted = normalizeBookmarks(bookmarks);
  const rows: Array<DocumentBookmark & { level: number }> = [];
  function visit(parentId: string | undefined, level: number) {
    sorted
      .filter((bookmark) => bookmark.parentId === parentId)
      .sort((a, b) => a.order - b.order)
      .forEach((bookmark) => {
        rows.push({ ...bookmark, level });
        if (bookmark.expanded) visit(bookmark.id, level + 1);
      });
  }
  visit(undefined, 1);
  return rows;
}

export function moveBookmark(bookmarks: DocumentBookmark[], bookmarkId: string, direction: -1 | 1) {
  const bookmark = bookmarks.find((item) => item.id === bookmarkId);
  if (!bookmark) return bookmarks;
  const siblings = bookmarks.filter((item) => item.parentId === bookmark.parentId).sort((a, b) => a.order - b.order);
  const index = siblings.findIndex((item) => item.id === bookmarkId);
  const targetIndex = index + direction;
  if (targetIndex < 0 || targetIndex >= siblings.length) return bookmarks;
  const swapped = siblings.map((item, itemIndex) => {
    if (itemIndex === index) return { ...item, order: targetIndex, updatedAt: new Date().toISOString() };
    if (itemIndex === targetIndex) return { ...item, order: index, updatedAt: new Date().toISOString() };
    return item;
  });
  const byId = new Map(swapped.map((item) => [item.id, item]));
  return normalizeBookmarks(bookmarks.map((item) => byId.get(item.id) ?? item));
}

export function nestBookmark(bookmarks: DocumentBookmark[], bookmarkId: string) {
  const bookmark = bookmarks.find((item) => item.id === bookmarkId);
  if (!bookmark) return bookmarks;
  const siblings = bookmarks.filter((item) => item.parentId === bookmark.parentId).sort((a, b) => a.order - b.order);
  const index = siblings.findIndex((item) => item.id === bookmarkId);
  const previous = siblings[index - 1];
  if (!previous || wouldCreateBookmarkCycle(bookmarks, bookmarkId, previous.id)) return bookmarks;
  return normalizeBookmarks(bookmarks.map((item) => (item.id === bookmarkId ? { ...item, parentId: previous.id, updatedAt: new Date().toISOString() } : item)));
}

export function outdentBookmark(bookmarks: DocumentBookmark[], bookmarkId: string) {
  const bookmark = bookmarks.find((item) => item.id === bookmarkId);
  if (!bookmark?.parentId) return bookmarks;
  const parent = bookmarks.find((item) => item.id === bookmark.parentId);
  return normalizeBookmarks(bookmarks.map((item) => (item.id === bookmarkId ? { ...item, parentId: parent?.parentId, updatedAt: new Date().toISOString() } : item)));
}

export function deleteBookmark(bookmarks: DocumentBookmark[], bookmarkId: string, mode: "deleteChildren" | "promoteChildren") {
  const descendantIds = new Set(getBookmarkDescendants(bookmarks, bookmarkId).map((bookmark) => bookmark.id));
  const deleted = bookmarks.find((bookmark) => bookmark.id === bookmarkId);
  if (!deleted) return bookmarks;
  if (mode === "deleteChildren") return normalizeBookmarks(bookmarks.filter((bookmark) => bookmark.id !== bookmarkId && !descendantIds.has(bookmark.id)));
  return normalizeBookmarks(bookmarks.filter((bookmark) => bookmark.id !== bookmarkId).map((bookmark) => (bookmark.parentId === bookmarkId ? { ...bookmark, parentId: deleted.parentId } : bookmark)));
}
