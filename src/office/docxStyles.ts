import { getAttr, getChildrenByLocalName, getElementsByLocalName, getFirstChildByLocalName, getXml, type DocxPackage } from "./docxPackage";
import type { DocxImportWarning, DocxTextRun } from "./types";

export type DocxParagraphFormat = {
  alignment?: "left" | "center" | "right";
  indentLeft?: number;
  indentRight?: number;
  firstLineIndent?: number;
  hangingIndent?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: number;
  keepWithNext?: boolean;
  keepLinesTogether?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
};

export type DocxRunFormat = Omit<Partial<DocxTextRun>, "id" | "text" | "hyperlink">;

export type DocxStyleDefinition = {
  id: string;
  type: "paragraph" | "character";
  name: string;
  basedOn?: string;
  isDefault?: boolean;
  headingLevel?: number;
  paragraph: DocxParagraphFormat;
  run: DocxRunFormat;
};

export type DocxStyleMap = {
  headingLevels: Map<string, number>;
  paragraphNames: Map<string, string>;
  paragraphStyles: Map<string, DocxStyleDefinition>;
  characterStyles: Map<string, DocxStyleDefinition>;
  defaultParagraphStyleId?: string;
  defaultCharacterStyleId?: string;
  documentDefaults: {
    paragraph: DocxParagraphFormat;
    run: DocxRunFormat;
  };
  resolveParagraphStyle: (styleId?: string) => DocxStyleDefinition | undefined;
  resolveCharacterStyle: (styleId?: string) => DocxStyleDefinition | undefined;
};

export type DocxNumberingLevel = {
  kind: "bullet" | "number" | "romanLower" | "romanUpper" | "alphaLower" | "alphaUpper";
  format?: string;
  text?: string;
  start: number;
  paragraph: DocxParagraphFormat;
};

export type DocxNumberingMap = Map<string, Map<string, DocxNumberingLevel>>;

export function parseDocxStyles(pkg: DocxPackage): DocxStyleMap {
  const xml = getXml(pkg, "word/styles.xml", "word/styles.xml");
  const headingLevels = new Map<string, number>();
  const paragraphNames = new Map<string, string>();
  const paragraphStyles = new Map<string, DocxStyleDefinition>();
  const characterStyles = new Map<string, DocxStyleDefinition>();
  const documentDefaults = parseDocumentDefaults(xml);
  let defaultParagraphStyleId: string | undefined;
  let defaultCharacterStyleId: string | undefined;

  if (xml) {
    for (const style of getElementsByLocalName(xml, "style")) {
      const type = getAttr(style, "type");
      if (type !== "paragraph" && type !== "character") continue;
      const styleId = getAttr(style, "styleId");
      if (!styleId) continue;
      const name = getAttr(getFirstChildByLocalName(style, "name"), "val") ?? styleId;
      const basedOn = getAttr(getFirstChildByLocalName(style, "basedOn"), "val");
      const isDefault = getAttr(style, "default") === "1" || getAttr(style, "default") === "true";
      const pPr = getFirstChildByLocalName(style, "pPr");
      const rPr = getFirstChildByLocalName(style, "rPr");
      const outlineLevel = getAttr(getFirstChildByLocalName(pPr ?? style, "outlineLvl"), "val");
      const headingLevel = outlineLevel !== undefined ? Number(outlineLevel) + 1 : getHeadingLevelFromName(name) ?? getHeadingLevelFromName(styleId);
      const definition: DocxStyleDefinition = {
        id: styleId,
        type,
        name,
        basedOn,
        isDefault,
        headingLevel: headingLevel && headingLevel >= 1 && headingLevel <= 9 ? headingLevel : undefined,
        paragraph: parseParagraphFormat(pPr),
        run: parseRunFormat(rPr),
      };
      if (type === "paragraph") {
        paragraphNames.set(styleId, name);
        paragraphStyles.set(styleId, definition);
        if (definition.headingLevel) headingLevels.set(styleId, definition.headingLevel);
        if (isDefault) defaultParagraphStyleId = styleId;
      } else {
        characterStyles.set(styleId, definition);
        if (isDefault) defaultCharacterStyleId = styleId;
      }
    }
  }

  function resolveStyle(styleId: string | undefined, map: Map<string, DocxStyleDefinition>, seen = new Set<string>()): DocxStyleDefinition | undefined {
    if (!styleId || seen.has(styleId)) return undefined;
    const style = map.get(styleId);
    if (!style) return undefined;
    seen.add(styleId);
    const base = resolveStyle(style.basedOn, map, seen);
    if (!base) return style;
    return {
      ...style,
      paragraph: { ...base.paragraph, ...style.paragraph },
      run: { ...base.run, ...style.run },
      headingLevel: style.headingLevel ?? base.headingLevel,
    };
  }

  return {
    headingLevels,
    paragraphNames,
    paragraphStyles,
    characterStyles,
    defaultParagraphStyleId,
    defaultCharacterStyleId,
    documentDefaults,
    resolveParagraphStyle: (styleId) => resolveStyle(styleId ?? defaultParagraphStyleId, paragraphStyles),
    resolveCharacterStyle: (styleId) => resolveStyle(styleId ?? defaultCharacterStyleId, characterStyles),
  };
}

export function parseDocxNumbering(pkg: DocxPackage): DocxNumberingMap {
  const xml = getXml(pkg, "word/numbering.xml", "word/numbering.xml");
  const map: DocxNumberingMap = new Map();
  if (!xml) return map;
  const abstractFormats = new Map<string, Map<string, DocxNumberingLevel>>();
  for (const abstract of getElementsByLocalName(xml, "abstractNum")) {
    const abstractId = getAttr(abstract, "abstractNumId");
    if (!abstractId) continue;
    const levels = new Map<string, DocxNumberingLevel>();
    for (const level of getChildrenByLocalName(abstract, "lvl")) {
      const ilvl = getAttr(level, "ilvl") ?? "0";
      const format = getAttr(getFirstChildByLocalName(level, "numFmt"), "val");
      levels.set(ilvl, {
        kind: getListKind(format),
        format,
        text: getAttr(getFirstChildByLocalName(level, "lvlText"), "val"),
        start: Number(getAttr(getFirstChildByLocalName(level, "start"), "val") ?? 1) || 1,
        paragraph: parseParagraphFormat(getFirstChildByLocalName(level, "pPr")),
      });
    }
    abstractFormats.set(abstractId, levels);
  }
  for (const num of getElementsByLocalName(xml, "num")) {
    const numId = getAttr(num, "numId");
    const abstractId = getAttr(getFirstChildByLocalName(num, "abstractNumId"), "val");
    if (numId && abstractId && abstractFormats.has(abstractId)) map.set(numId, abstractFormats.get(abstractId)!);
  }
  return map;
}

export function mergeParagraphFormat(...formats: Array<DocxParagraphFormat | undefined>) {
  return Object.assign({}, ...formats.filter(Boolean));
}

export function mergeRunFormat(...formats: Array<DocxRunFormat | undefined>) {
  return Object.assign({}, ...formats.filter(Boolean));
}

export function parseParagraphFormat(pPr: Element | null | undefined): DocxParagraphFormat {
  if (!pPr) return {};
  const indent = getFirstChildByLocalName(pPr, "ind");
  const spacing = getFirstChildByLocalName(pPr, "spacing");
  const lineValue = getAttr(spacing, "line");
  return {
    alignment: parseAlignment(getAttr(getFirstChildByLocalName(pPr, "jc"), "val")),
    indentLeft: twipsToPoints(getAttr(indent, "left")),
    indentRight: twipsToPoints(getAttr(indent, "right")),
    firstLineIndent: twipsToPoints(getAttr(indent, "firstLine")),
    hangingIndent: twipsToPoints(getAttr(indent, "hanging")),
    spacingBefore: twipsToPoints(getAttr(spacing, "before")),
    spacingAfter: twipsToPoints(getAttr(spacing, "after")),
    lineSpacing: lineValue ? Math.max(0.8, Number(lineValue) / 240) : undefined,
    keepWithNext: hasEnabledFlag(pPr, "keepNext"),
    keepLinesTogether: hasEnabledFlag(pPr, "keepLines"),
    pageBreakBefore: hasEnabledFlag(pPr, "pageBreakBefore"),
    widowControl: hasEnabledFlag(pPr, "widowControl"),
  };
}

export function parseRunFormat(rPr: Element | null | undefined): DocxRunFormat {
  if (!rPr) return {};
  const fontSize = halfPointsToPoints(getAttr(getFirstChildByLocalName(rPr, "sz"), "val"));
  const color = normalizeColor(getAttr(getFirstChildByLocalName(rPr, "color"), "val"));
  const highlight = normalizeHighlight(getAttr(getFirstChildByLocalName(rPr, "highlight"), "val"));
  const vertical = getAttr(getFirstChildByLocalName(rPr, "vertAlign"), "val");
  return {
    bold: hasEnabledFlag(rPr, "b"),
    italic: hasEnabledFlag(rPr, "i"),
    underline: Boolean(getFirstChildByLocalName(rPr, "u")),
    strike: hasEnabledFlag(rPr, "strike") || hasEnabledFlag(rPr, "dstrike"),
    color,
    backgroundColor: highlight,
    fontSize,
    fontFamily: getAttr(getFirstChildByLocalName(rPr, "rFonts"), "ascii") ?? getAttr(getFirstChildByLocalName(rPr, "rFonts"), "hAnsi"),
    letterSpacing: twipsToPoints(getAttr(getFirstChildByLocalName(rPr, "spacing"), "val")),
    verticalAlign: vertical === "superscript" || vertical === "subscript" ? vertical : undefined,
    smallCaps: hasEnabledFlag(rPr, "smallCaps"),
    allCaps: hasEnabledFlag(rPr, "caps"),
  };
}

export function collectStyleWarnings(styles: DocxStyleMap, warnings: DocxImportWarning[]) {
  const fonts = new Set<string>();
  for (const style of [...styles.paragraphStyles.values(), ...styles.characterStyles.values()]) {
    if (style.run.fontFamily) fonts.add(style.run.fontFamily);
  }
  for (const font of fonts) {
    if (!isBuiltInFallbackFont(font)) warnings.push({ code: "font-substitution", message: `${font} is referenced by Word styles and will use the selected PDF fallback font unless available in a later desktop conversion path.` });
  }
}

function parseDocumentDefaults(xml: Document | null) {
  const docDefaults = xml ? getElementsByLocalName(xml, "docDefaults")[0] : undefined;
  const pPrDefault = docDefaults ? getElementsByLocalName(docDefaults, "pPrDefault")[0] : undefined;
  const rPrDefault = docDefaults ? getElementsByLocalName(docDefaults, "rPrDefault")[0] : undefined;
  return {
    paragraph: parseParagraphFormat(pPrDefault ? getFirstChildByLocalName(pPrDefault, "pPr") : undefined),
    run: parseRunFormat(rPrDefault ? getFirstChildByLocalName(rPrDefault, "rPr") : undefined),
  };
}

function getListKind(format?: string): DocxNumberingLevel["kind"] {
  if (format === "bullet") return "bullet";
  if (format === "lowerRoman") return "romanLower";
  if (format === "upperRoman") return "romanUpper";
  if (format === "lowerLetter") return "alphaLower";
  if (format === "upperLetter") return "alphaUpper";
  return "number";
}

function getHeadingLevelFromName(value?: string) {
  if (!value) return undefined;
  const normalized = value.toLowerCase().replace(/\s+/g, "");
  const match = normalized.match(/^heading([1-9])$/);
  return match ? Number(match[1]) : undefined;
}

function parseAlignment(value?: string): DocxParagraphFormat["alignment"] {
  if (value === "center") return "center";
  if (value === "right" || value === "end") return "right";
  return value === "both" ? "left" : undefined;
}

function hasEnabledFlag(parent: Element, localName: string) {
  const element = getFirstChildByLocalName(parent, localName);
  if (!element) return undefined;
  const value = getAttr(element, "val");
  return value === undefined || (value !== "0" && value !== "false");
}

function twipsToPoints(value?: string) {
  if (!value) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed / 20 : undefined;
}

function halfPointsToPoints(value?: string) {
  if (!value) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed / 2 : undefined;
}

function normalizeColor(value?: string) {
  if (!value || value === "auto") return undefined;
  return `#${value.replace("#", "").padStart(6, "0").slice(0, 6)}`;
}

function normalizeHighlight(value?: string) {
  if (!value || value === "none") return undefined;
  const colors: Record<string, string> = {
    yellow: "#fff59d",
    green: "#bbf7d0",
    cyan: "#bae6fd",
    magenta: "#f0abfc",
    blue: "#bfdbfe",
    red: "#fecaca",
    darkBlue: "#1d4ed8",
    darkCyan: "#0e7490",
    darkGreen: "#166534",
    darkMagenta: "#86198f",
    darkRed: "#991b1b",
    darkYellow: "#a16207",
    black: "#111827",
  };
  return colors[value] ?? undefined;
}

function isBuiltInFallbackFont(font: string) {
  return ["helvetica", "arial", "times", "times new roman", "courier", "courier new"].includes(font.toLowerCase());
}
