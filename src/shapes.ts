export type ShapeKind = "rectangle" | "roundedRectangle" | "ellipse" | "line" | "arrow" | "doubleArrow" | "polygon" | "cloud" | "callout";
export type ShapeDashStyle = "solid" | "dashed" | "dotted";
export type ShapeArrowhead = "none" | "arrow";

export type ShapeStyle = {
  shapeKind: ShapeKind;
  strokeColor: string;
  fillColor: string;
  strokeWidth: number;
  strokeOpacity: number;
  fillOpacity: number;
  dashStyle: ShapeDashStyle;
  startArrowhead: ShapeArrowhead;
  endArrowhead: ShapeArrowhead;
  cornerRadius: number;
  calloutText: string;
  leaderEnd?: { x: number; y: number };
};

export const shapeToolLabels: Record<ShapeKind, string> = {
  rectangle: "Rectangle",
  roundedRectangle: "Rounded rectangle",
  ellipse: "Ellipse",
  line: "Line",
  arrow: "Arrow",
  doubleArrow: "Double-ended arrow",
  polygon: "Polygon",
  cloud: "Cloud",
  callout: "Text callout",
};

export const shapeDashStyles: ShapeDashStyle[] = ["solid", "dashed", "dotted"];

export function createShapeStyle(shapeKind: ShapeKind, patch: Partial<ShapeStyle> = {}): ShapeStyle {
  const isLine = shapeKind === "line" || shapeKind === "arrow" || shapeKind === "doubleArrow";
  return {
    shapeKind,
    strokeColor: "#d8342a",
    fillColor: "#facc15",
    strokeWidth: 2,
    strokeOpacity: 1,
    fillOpacity: isLine ? 0 : 0.18,
    dashStyle: "solid",
    startArrowhead: shapeKind === "doubleArrow" ? "arrow" : "none",
    endArrowhead: shapeKind === "arrow" || shapeKind === "doubleArrow" ? "arrow" : "none",
    cornerRadius: shapeKind === "roundedRectangle" ? 12 : 0,
    calloutText: shapeKind === "callout" ? "Callout" : "",
    leaderEnd: shapeKind === "callout" ? { x: 48, y: 96 } : undefined,
    ...patch,
  };
}

export function isShapeKind(value: string): value is ShapeKind {
  return value in shapeToolLabels;
}

export function getDashArray(style: ShapeDashStyle, width: number) {
  if (style === "dashed") return `${Math.max(4, width * 3)} ${Math.max(3, width * 2)}`;
  if (style === "dotted") return `${Math.max(1, width)} ${Math.max(3, width * 2)}`;
  return undefined;
}

export function getPdfDashArray(style: ShapeDashStyle, width: number) {
  if (style === "dashed") return [Math.max(4, width * 3), Math.max(3, width * 2)];
  if (style === "dotted") return [Math.max(1, width), Math.max(3, width * 2)];
  return undefined;
}
