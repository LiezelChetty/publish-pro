import type { PublishingNumberFormat, PublishingTokenContext } from "./types";

export function renderPublishingTokens(template: string, context: PublishingTokenContext) {
  return [
    ["{page}", String(context.pageNumber)],
    ["{pages}", String(context.pageCount)],
    ["{label}", context.pageLabel || String(context.pageNumber)],
    ["{project}", context.projectName],
    ["{client}", context.client],
    ["{date}", context.date],
    ["{filename}", context.filename],
  ].reduce((output, [token, value]) => output.split(token).join(value), template);
}

export function formatPageNumber(format: PublishingNumberFormat, value: number, pageCount: number, customTemplate: string) {
  if (format === "decimal2") return String(value).padStart(2, "0");
  if (format === "decimal3") return String(value).padStart(3, "0");
  if (format === "romanLower") return toRoman(value).toLowerCase();
  if (format === "romanUpper") return toRoman(value);
  if (format === "alphaLower") return toAlpha(value).toLowerCase();
  if (format === "alphaUpper") return toAlpha(value);
  if (format === "page") return `Page ${value}`;
  if (format === "pageOfPages") return `Page ${value} of ${pageCount}`;
  if (format === "custom") return customTemplate || "{page}";
  return String(value);
}

export function formatNumberingTemplate({
  format,
  value,
  documentPages,
  numberedPages,
  sectionPages,
  customTemplate,
  prefix = "",
  suffix = "",
}: {
  format: PublishingNumberFormat;
  value: number;
  documentPages: number;
  numberedPages: number;
  sectionPages: number;
  customTemplate: string;
  prefix?: string;
  suffix?: string;
}) {
  if (format === "pageOfPages") {
    const current = `${prefix}${value}${suffix}`;
    const total = `${prefix}${sectionPages}${suffix}`;
    return `Page ${current} of ${total}`;
  }
  if (format !== "custom") {
    return `${prefix}${formatPageNumber(format, value, sectionPages, customTemplate)}${suffix}`;
  }
  const template = customTemplate || "{page}";
  const current = `${prefix}${value}${suffix}`;
  const sectionTotal = `${prefix}${sectionPages}${suffix}`;
  return [
    ["{page}", current],
    ["{pages}", String(documentPages)],
    ["{numberedPages}", String(numberedPages)],
    ["{sectionPages}", sectionTotal],
  ].reduce((output, [token, replacement]) => output.split(token).join(replacement), template);
}

function toRoman(value: number) {
  const pairs: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];
  let remaining = Math.max(1, Math.floor(value));
  let output = "";
  for (const [amount, symbol] of pairs) {
    while (remaining >= amount) {
      output += symbol;
      remaining -= amount;
    }
  }
  return output;
}

function toAlpha(value: number) {
  let n = Math.max(1, Math.floor(value));
  let output = "";
  while (n > 0) {
    n -= 1;
    output = String.fromCharCode(65 + (n % 26)) + output;
    n = Math.floor(n / 26);
  }
  return output;
}
