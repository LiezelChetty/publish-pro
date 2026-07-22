import { PDFArray, PDFDict, PDFHexString, PDFName, PDFNumber, type PDFDocument, type PDFPage, type PDFRef } from "pdf-lib";
import { getBookmarkLevel } from "./bookmarks";
import type { DocumentBookmark, TocLine } from "./types";

export function addInternalLinkAnnotation(pdf: PDFDocument, page: PDFPage, targetPage: PDFPage, rect: { x: number; y: number; width: number; height: number }) {
  const context = pdf.context;
  const destination = context.obj([targetPage.ref, PDFName.of("XYZ"), PDFNumber.of(0), PDFNumber.of(targetPage.getHeight()), PDFNumber.of(0)]);
  const annotation = context.register(
    context.obj({
      Type: "Annot",
      Subtype: "Link",
      Rect: [rect.x, rect.y, rect.x + rect.width, rect.y + rect.height],
      Border: [0, 0, 0],
      Dest: destination,
    })
  );
  const existing = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
  const annots = existing ?? context.obj([]);
  annots.push(annotation);
  page.node.set(PDFName.of("Annots"), annots);
}

export function addTocLinks(pdf: PDFDocument, tocPages: PDFPage[], allPages: PDFPage[], tocLines: TocLine[], pageIdToExportIndex: Map<string, number>) {
  for (const line of tocLines) {
    const tocPage = tocPages[line.pageIndex];
    const targetIndex = pageIdToExportIndex.get(line.pageId);
    const targetPage = targetIndex === undefined ? undefined : allPages[targetIndex];
    if (!tocPage || !targetPage) continue;
    addInternalLinkAnnotation(pdf, tocPage, targetPage, {
      x: line.x,
      y: line.y - 2,
      width: line.width,
      height: line.height,
    });
    if (line.pageText) {
      addInternalLinkAnnotation(pdf, tocPage, targetPage, {
        x: Math.max(line.x, tocPage.getWidth() - 96),
        y: line.y - 2,
        width: 88,
        height: line.height,
      });
    }
  }
}

export function addPdfOutline(pdf: PDFDocument, bookmarks: DocumentBookmark[], pageIdToExportIndex: Map<string, number>) {
  const validBookmarks = bookmarks
    .filter((bookmark) => pageIdToExportIndex.has(bookmark.pageId))
    .sort((a, b) => getBookmarkPath(bookmarks, a).localeCompare(getBookmarkPath(bookmarks, b)));
  if (validBookmarks.length === 0) return;

  const context = pdf.context;
  const pages = pdf.getPages();
  const outlineRoot = context.obj({ Type: "Outlines" });
  const rootRef = context.register(outlineRoot);
  const refs = new Map<string, ReturnType<typeof context.register>>();
  for (const bookmark of validBookmarks) refs.set(bookmark.id, context.register(context.obj({})));

  for (const bookmark of validBookmarks) {
    const ref = refs.get(bookmark.id);
    if (!ref) continue;
    const outline = context.lookup(ref, PDFDict);
    const targetPage = pages[pageIdToExportIndex.get(bookmark.pageId) ?? 0];
    const parentRef = getOutlineParentRef(bookmark, refs, rootRef);
    outline.set(PDFName.of("Title"), PDFHexString.fromText(bookmark.title || "Untitled"));
    outline.set(PDFName.of("Parent"), parentRef);
    outline.set(PDFName.of("Dest"), context.obj([targetPage.ref, PDFName.of("XYZ"), PDFNumber.of(0), PDFNumber.of(targetPage.getHeight()), PDFNumber.of(0)]));
    const parentId = getValidParentId(bookmark, refs);
    const siblings = validBookmarks.filter((item) => getValidParentId(item, refs) === parentId).sort((a, b) => a.order - b.order);
    const index = siblings.findIndex((item) => item.id === bookmark.id);
    if (siblings[index - 1]) outline.set(PDFName.of("Prev"), refs.get(siblings[index - 1].id)!);
    if (siblings[index + 1]) outline.set(PDFName.of("Next"), refs.get(siblings[index + 1].id)!);
    const children = validBookmarks.filter((item) => getValidParentId(item, refs) === bookmark.id).sort((a, b) => a.order - b.order);
    if (children.length > 0) {
      outline.set(PDFName.of("First"), refs.get(children[0].id)!);
      outline.set(PDFName.of("Last"), refs.get(children[children.length - 1].id)!);
      const descendantCount = countValidDescendants(validBookmarks, bookmark.id, refs);
      outline.set(PDFName.of("Count"), PDFNumber.of(bookmark.expanded ? descendantCount : -descendantCount));
    }
  }

  for (const bookmark of validBookmarks) {
    const outline = context.lookup(refs.get(bookmark.id)!, PDFDict);
    outline.set(PDFName.of("Parent"), getOutlineParentRef(bookmark, refs, rootRef));
  }
  const topLevel = validBookmarks.filter((bookmark) => !getValidParentId(bookmark, refs)).sort((a, b) => a.order - b.order);
  outlineRoot.set(PDFName.of("First"), refs.get(topLevel[0].id)!);
  outlineRoot.set(PDFName.of("Last"), refs.get(topLevel[topLevel.length - 1].id)!);
  outlineRoot.set(PDFName.of("Count"), PDFNumber.of(validBookmarks.length));
  pdf.catalog.set(PDFName.of("Outlines"), rootRef);
  pdf.catalog.set(PDFName.of("PageMode"), PDFName.of("UseOutlines"));
}

function getValidParentId(bookmark: DocumentBookmark, refs: Map<string, unknown>) {
  return bookmark.parentId && refs.has(bookmark.parentId) ? bookmark.parentId : undefined;
}

function getOutlineParentRef(bookmark: DocumentBookmark, refs: Map<string, PDFRef>, rootRef: PDFRef) {
  const parentId = getValidParentId(bookmark, refs);
  return parentId ? refs.get(parentId)! : rootRef;
}

function countValidDescendants(bookmarks: DocumentBookmark[], bookmarkId: string, refs: Map<string, unknown>): number {
  const children = bookmarks.filter((bookmark) => getValidParentId(bookmark, refs) === bookmarkId);
  return children.reduce((count, child) => count + 1 + countValidDescendants(bookmarks, child.id, refs), 0);
}

function getBookmarkPath(bookmarks: DocumentBookmark[], bookmark: DocumentBookmark) {
  const path = [`${bookmark.order}`.padStart(5, "0")];
  let parent = bookmarks.find((item) => item.id === bookmark.parentId);
  while (parent) {
    path.unshift(`${parent.order}`.padStart(5, "0"));
    parent = bookmarks.find((item) => item.id === parent?.parentId);
  }
  return `${path.join(".")}.${getBookmarkLevel(bookmarks, bookmark)}`;
}
