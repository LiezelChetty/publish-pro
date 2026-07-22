import type { DocumentBookmark } from "../navigation/types";
import type { DocxHeading } from "./types";

export function createBookmarksFromDocxHeadings(headings: DocxHeading[], pageIdsByNumber: Map<number, string>): DocumentBookmark[] {
  const now = new Date().toISOString();
  const bookmarks: DocumentBookmark[] = [];
  const parentByLevel = new Map<number, string>();
  const orderByParent = new Map<string, number>();
  for (const heading of headings) {
    const title = heading.title.trim();
    const pageId = pageIdsByNumber.get(heading.pageNumber);
    if (!title || !pageId) continue;
    const level = Math.min(Math.max(heading.level, 1), 9);
    const parentId = level > 1 ? findParentId(parentByLevel, level) : undefined;
    const orderKey = parentId ?? "root";
    const order = orderByParent.get(orderKey) ?? 0;
    const bookmark: DocumentBookmark = {
      id: crypto.randomUUID(),
      title,
      pageId,
      parentId,
      order,
      expanded: true,
      createdAt: now,
      updatedAt: now,
    };
    bookmarks.push(bookmark);
    parentByLevel.set(level, bookmark.id);
    for (const existingLevel of Array.from(parentByLevel.keys())) {
      if (existingLevel > level) parentByLevel.delete(existingLevel);
    }
    orderByParent.set(orderKey, order + 1);
  }
  return bookmarks;
}

function findParentId(parentByLevel: Map<number, string>, level: number) {
  for (let candidate = level - 1; candidate >= 1; candidate -= 1) {
    const parentId = parentByLevel.get(candidate);
    if (parentId) return parentId;
  }
  return undefined;
}
