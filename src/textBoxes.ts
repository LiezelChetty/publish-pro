import { StandardFonts, type PDFFont } from "pdf-lib";

export type TextBoxFontFamily = "Helvetica" | "Times Roman" | "Courier";
export type TextBoxAlign = "left" | "center" | "right";
export type TextBoxPreset = "body" | "heading" | "caption" | "redaction" | "approved" | "draft";

export type TextBoxStyle = {
  fontFamily: TextBoxFontFamily;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  align: TextBoxAlign;
  lineHeight: number;
  letterSpacing: number;
  backgroundColor: string;
  backgroundOpacity: number;
  borderColor: string;
  borderWidth: number;
  padding: number;
};

export const textBoxFontFamilies: TextBoxFontFamily[] = ["Helvetica", "Times Roman", "Courier"];
export const textBoxAlignments: TextBoxAlign[] = ["left", "center", "right"];

export const defaultTextBoxStyle: TextBoxStyle = {
  fontFamily: "Helvetica",
  bold: false,
  italic: false,
  underline: false,
  align: "left",
  lineHeight: 1.25,
  letterSpacing: 0,
  backgroundColor: "#ffffff",
  backgroundOpacity: 0,
  borderColor: "#111827",
  borderWidth: 0,
  padding: 8,
};

export const textBoxPresets: Array<{ label: string; value: TextBoxPreset }> = [
  { label: "Body Text", value: "body" },
  { label: "Heading", value: "heading" },
  { label: "Caption", value: "caption" },
  { label: "Redaction Label", value: "redaction" },
  { label: "Approved", value: "approved" },
  { label: "Draft", value: "draft" },
];

export function createDefaultTextBoxStyle(patch: Partial<TextBoxStyle> = {}): TextBoxStyle {
  return { ...defaultTextBoxStyle, ...patch };
}

export function getPresetTextStyle(preset: TextBoxPreset): Partial<TextBoxStyle> & { text?: string; color?: string; size?: number } {
  if (preset === "heading") return { size: 20, bold: true, lineHeight: 1.18 };
  if (preset === "caption") return { size: 10, color: "#475569", lineHeight: 1.2 };
  if (preset === "redaction") {
    return {
      text: "REDACTED",
      color: "#ffffff",
      size: 12,
      bold: true,
      align: "center",
      backgroundColor: "#111827",
      backgroundOpacity: 1,
      padding: 6,
    };
  }
  if (preset === "approved") return { text: "APPROVED", color: "#16a34a", size: 16, bold: true, align: "center" };
  if (preset === "draft") return { text: "DRAFT", color: "#d8342a", size: 18, bold: true, align: "center", letterSpacing: 1 };
  return { size: 12, color: "#111827", lineHeight: 1.25 };
}

export function getStandardFontName(style: TextBoxStyle) {
  if (style.fontFamily === "Times Roman") {
    if (style.bold && style.italic) return StandardFonts.TimesRomanBoldItalic;
    if (style.bold) return StandardFonts.TimesRomanBold;
    if (style.italic) return StandardFonts.TimesRomanItalic;
    return StandardFonts.TimesRoman;
  }

  if (style.fontFamily === "Courier") {
    if (style.bold && style.italic) return StandardFonts.CourierBoldOblique;
    if (style.bold) return StandardFonts.CourierBold;
    if (style.italic) return StandardFonts.CourierOblique;
    return StandardFonts.Courier;
  }

  if (style.bold && style.italic) return StandardFonts.HelveticaBoldOblique;
  if (style.bold) return StandardFonts.HelveticaBold;
  if (style.italic) return StandardFonts.HelveticaOblique;
  return StandardFonts.Helvetica;
}

export function getCssFontFamily(fontFamily: TextBoxFontFamily) {
  if (fontFamily === "Times Roman") return "\"Times New Roman\", Times, serif";
  if (fontFamily === "Courier") return "\"Courier New\", Courier, monospace";
  return "Helvetica, Arial, sans-serif";
}

export function wrapTextForBox(text: string, maxWidth: number, fontSize: number, font: Pick<PDFFont, "widthOfTextAtSize"> | null, letterSpacing = 0) {
  const hardLines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const lines: string[] = [];

  for (const hardLine of hardLines) {
    if (!hardLine) {
      lines.push("");
      continue;
    }

    const words = hardLine.split(/(\s+)/).filter((part) => part.length > 0);
    let line = "";
    for (const word of words) {
      const next = line ? `${line}${word}` : word.trimStart();
      if (line && measureTextWidth(next, fontSize, font, letterSpacing) > maxWidth) {
        lines.push(line.trimEnd());
        line = word.trimStart();
      } else {
        line = next;
      }
    }

    if (line) {
      while (measureTextWidth(line, fontSize, font, letterSpacing) > maxWidth && line.length > 1) {
        const splitAt = findFittingBreak(line, maxWidth, fontSize, font, letterSpacing);
        lines.push(line.slice(0, splitAt));
        line = line.slice(splitAt);
      }
      lines.push(line);
    }
  }

  return lines;
}

export function measureTextWidth(text: string, fontSize: number, font: Pick<PDFFont, "widthOfTextAtSize"> | null, letterSpacing = 0) {
  const baseWidth = font ? font.widthOfTextAtSize(text, fontSize) : text.length * fontSize * 0.56;
  return baseWidth + Math.max(0, text.length - 1) * letterSpacing;
}

function findFittingBreak(text: string, maxWidth: number, fontSize: number, font: Pick<PDFFont, "widthOfTextAtSize"> | null, letterSpacing: number) {
  for (let index = Math.max(1, text.length - 1); index > 1; index -= 1) {
    if (measureTextWidth(text.slice(0, index), fontSize, font, letterSpacing) <= maxWidth) return index;
  }
  return 1;
}
