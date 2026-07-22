import { getAttr, getChildrenByLocalName, getElementsByLocalName, getFirstChildByLocalName, getXml, type DocxPackage } from "./docxPackage";

export type DocxStyleMap = {
  headingLevels: Map<string, number>;
  paragraphNames: Map<string, string>;
};

export type DocxNumberingMap = Map<string, Map<string, "bullet" | "number">>;

export function parseDocxStyles(pkg: DocxPackage): DocxStyleMap {
  const xml = getXml(pkg, "word/styles.xml", "word/styles.xml");
  const headingLevels = new Map<string, number>();
  const paragraphNames = new Map<string, string>();
  if (!xml) return { headingLevels, paragraphNames };

  for (const style of getElementsByLocalName(xml, "style")) {
    if (getAttr(style, "type") !== "paragraph") continue;
    const styleId = getAttr(style, "styleId");
    if (!styleId) continue;
    const name = getAttr(getFirstChildByLocalName(style, "name"), "val") ?? styleId;
    paragraphNames.set(styleId, name);
    const outlineLevel = getAttr(getFirstChildByLocalName(getFirstChildByLocalName(style, "pPr") ?? style, "outlineLvl"), "val");
    const styleLevel = getHeadingLevelFromName(name) ?? getHeadingLevelFromName(styleId);
    const level = outlineLevel !== undefined ? Number(outlineLevel) + 1 : styleLevel;
    if (level && level >= 1 && level <= 9) headingLevels.set(styleId, level);
  }
  return { headingLevels, paragraphNames };
}

export function parseDocxNumbering(pkg: DocxPackage): DocxNumberingMap {
  const xml = getXml(pkg, "word/numbering.xml", "word/numbering.xml");
  const map: DocxNumberingMap = new Map();
  if (!xml) return map;
  const abstractFormats = new Map<string, Map<string, "bullet" | "number">>();
  for (const abstract of getElementsByLocalName(xml, "abstractNum")) {
    const abstractId = getAttr(abstract, "abstractNumId");
    if (!abstractId) continue;
    const levels = new Map<string, "bullet" | "number">();
    for (const level of getChildrenByLocalName(abstract, "lvl")) {
      const ilvl = getAttr(level, "ilvl") ?? "0";
      const format = getAttr(getFirstChildByLocalName(level, "numFmt"), "val");
      levels.set(ilvl, format === "bullet" ? "bullet" : "number");
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

function getHeadingLevelFromName(value?: string) {
  if (!value) return undefined;
  const normalized = value.toLowerCase().replace(/\s+/g, "");
  const match = normalized.match(/^heading([1-9])$/);
  return match ? Number(match[1]) : undefined;
}
