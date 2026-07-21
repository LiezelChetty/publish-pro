export type TextMarkupKind = "textHighlight" | "underline" | "strikethrough";

export type MarkupRect = {
  x: number;
  y: number;
  width: number;
  height: number;
};

export type PageBounds = {
  width: number;
  height: number;
};

export function isTextMarkupKind(kind: string): kind is TextMarkupKind {
  return kind === "textHighlight" || kind === "underline" || kind === "strikethrough";
}

export function getMarkupBounds(rects: MarkupRect[]) {
  if (rects.length === 0) return null;

  const minX = Math.min(...rects.map((rect) => rect.x));
  const minY = Math.min(...rects.map((rect) => rect.y));
  const maxX = Math.max(...rects.map((rect) => rect.x + rect.width));
  const maxY = Math.max(...rects.map((rect) => rect.y + rect.height));

  return {
    x: minX,
    y: minY,
    width: Math.max(1, maxX - minX),
    height: Math.max(1, maxY - minY),
  };
}

export function normalizeSelectionRects(clientRects: DOMRectList, pageRect: DOMRect, zoom: number, page: PageBounds) {
  const rects = Array.from(clientRects)
    .map((rect) => ({
      x: clamp((rect.left - pageRect.left) / zoom, 0, page.width),
      y: clamp((rect.top - pageRect.top) / zoom, 0, page.height),
      width: Math.min(rect.width / zoom, page.width),
      height: Math.min(rect.height / zoom, page.height),
    }))
    .map((rect) => ({
      ...rect,
      width: Math.max(0, Math.min(rect.width, page.width - rect.x)),
      height: Math.max(0, Math.min(rect.height, page.height - rect.y)),
    }))
    .filter((rect) => rect.width > 3 && rect.height > 3);

  return mergeLineRects(rects);
}

export function moveMarkupRects(rects: MarkupRect[] | undefined, dx: number, dy: number, page: PageBounds) {
  if (!rects || rects.length === 0) return rects;

  const moved = rects.map((rect) => ({ ...rect, x: rect.x + dx, y: rect.y + dy }));
  const bounds = getMarkupBounds(moved);
  if (!bounds) return moved;

  const correctionX = bounds.x < 0 ? -bounds.x : bounds.x + bounds.width > page.width ? page.width - (bounds.x + bounds.width) : 0;
  const correctionY = bounds.y < 0 ? -bounds.y : bounds.y + bounds.height > page.height ? page.height - (bounds.y + bounds.height) : 0;

  return moved.map((rect) => ({ ...rect, x: rect.x + correctionX, y: rect.y + correctionY }));
}

export function scaleMarkupRect(rect: MarkupRect, sourcePage: PageBounds, targetPage: PageBounds) {
  return {
    x: (rect.x / sourcePage.width) * targetPage.width,
    y: targetPage.height - ((rect.y + rect.height) / sourcePage.height) * targetPage.height,
    width: (rect.width / sourcePage.width) * targetPage.width,
    height: (rect.height / sourcePage.height) * targetPage.height,
  };
}

function mergeLineRects(rects: MarkupRect[]) {
  const sorted = [...rects].sort((left, right) => (Math.abs(left.y - right.y) > 3 ? left.y - right.y : left.x - right.x));
  const merged: MarkupRect[] = [];

  for (const rect of sorted) {
    const previous = merged[merged.length - 1];
    const sameLine = previous && Math.abs(previous.y - rect.y) < Math.max(3, Math.min(previous.height, rect.height) * 0.45);
    const closeHorizontally = previous && rect.x <= previous.x + previous.width + 4;

    if (sameLine && closeHorizontally) {
      const minX = Math.min(previous.x, rect.x);
      const minY = Math.min(previous.y, rect.y);
      const maxX = Math.max(previous.x + previous.width, rect.x + rect.width);
      const maxY = Math.max(previous.y + previous.height, rect.y + rect.height);
      merged[merged.length - 1] = { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
    } else {
      merged.push(rect);
    }
  }

  return merged;
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}
