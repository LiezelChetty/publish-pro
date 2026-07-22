#!/usr/bin/env node
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { basename, join } from "node:path";
import { spawnSync } from "node:child_process";
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFRef,
  PDFString,
  StandardFonts,
  degrees,
  rgb,
} from "pdf-lib";

const args = process.argv.slice(2);

if (args.includes("--generate-fixture")) {
  const fixtureDir = "/private/tmp/publish-pro-validation";
  mkdirSync(fixtureDir, { recursive: true });
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const pages = [
    ["Cover", [612, 792], 0],
    ["TOC i", [612, 792], 0],
    ["TOC ii", [612, 792], 0],
    ["TOC iii", [612, 792], 0],
    ["TOC iv", [612, 792], 0],
    ["Main 1", [612, 792], 0],
    ["Main 2", [612, 792], 0],
    ["Main 3 landscape", [841.89, 595.28], 0],
    ["Main 4", [612, 792], 0],
    ["Main 5 rotated", [612, 792], 90],
    ["Main 6", [612, 792], 0],
    ["Main 7", [612, 792], 0],
    ["Main 8", [612, 792], 0],
    ["Appendix A-1", [595.28, 841.89], 0],
    ["Appendix A-2 mixed legal", [612, 1008], 0],
    ["Appendix A-3", [612, 792], 0],
  ];
  for (const [title, size, rotation] of pages) {
    const page = pdf.addPage(size);
    if (rotation) page.setRotation(degrees(rotation));
    page.drawText(`Publish Pro navigation fixture - ${title}`, { x: 54, y: size[1] - 72, size: 16, font, color: rgb(0.08, 0.1, 0.16) });
    page.drawText("Use this generated source PDF to build a repeatable bookmark, TOC, and numbering validation project.", { x: 54, y: size[1] - 104, size: 10, font });
  }
  const output = join(fixtureDir, "publish-pro-navigation-source.pdf");
  writeFileSync(output, await pdf.save());
  console.log(output);
  process.exit(0);
}

const filePath = args[0];
if (!filePath) {
  console.error("Usage: node scripts/validate-navigation-export.mjs <pdf-path> [--pages N] [--min-outlines N] [--min-links N] [--expect-text TEXT]\n       node scripts/validate-navigation-export.mjs --generate-fixture");
  process.exit(1);
}

const options = readOptions(args.slice(1));
if (!existsSync(filePath)) fail(`PDF does not exist: ${filePath}`);
const stats = statSync(filePath);
if (stats.size <= 0) fail(`PDF is empty: ${filePath}`);

const bytes = readFileSync(filePath);
const pdf = await PDFDocument.load(bytes);
const pages = pdf.getPages();
if (options.pages !== undefined && pages.length !== options.pages) fail(`Expected ${options.pages} pages, found ${pages.length}.`);

const pageRefs = new Set(pages.map((page) => refKey(page.ref)));
const outlineReport = inspectOutlines(pdf, pageRefs);
const linkReport = inspectLinkAnnotations(pdf, pages, pageRefs);
if (options.minOutlines !== undefined && outlineReport.nodeCount < options.minOutlines) fail(`Expected at least ${options.minOutlines} outline entries, found ${outlineReport.nodeCount}.`);
if (options.minLinks !== undefined && linkReport.linkCount < options.minLinks) fail(`Expected at least ${options.minLinks} link annotations, found ${linkReport.linkCount}.`);
if (outlineReport.errors.length > 0) fail(`Outline validation failed:\n${outlineReport.errors.join("\n")}`);
if (linkReport.errors.length > 0) fail(`Link annotation validation failed:\n${linkReport.errors.join("\n")}`);

const textReport = inspectExpectedText(filePath, options.expectedText);
if (textReport.errors.length > 0) fail(textReport.errors.join("\n"));

const renderReport = renderRepresentativePages(filePath, pages.length);
const dimensions = pages.map((page, index) => {
  const size = page.getSize();
  return {
    page: index + 1,
    width: Number(size.width.toFixed(2)),
    height: Number(size.height.toFixed(2)),
    rotation: page.getRotation().angle,
  };
});

console.log(JSON.stringify({
  filePath,
  size: stats.size,
  pageCount: pages.length,
  dimensions,
  outlines: outlineReport,
  links: linkReport,
  text: textReport,
  render: renderReport,
}, null, 2));

function readOptions(values) {
  const options = { expectedText: [] };
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (value === "--pages") options.pages = Number(values[++index]);
    if (value === "--min-outlines") options.minOutlines = Number(values[++index]);
    if (value === "--min-links") options.minLinks = Number(values[++index]);
    if (value === "--expect-text") options.expectedText.push(values[++index]);
  }
  return options;
}

function inspectOutlines(pdf, pageRefs) {
  const context = pdf.context;
  const outlineRef = pdf.catalog.get(PDFName.of("Outlines"));
  const report = { nodeCount: 0, maxDepth: 0, titles: [], errors: [] };
  if (!outlineRef) return report;
  const root = lookupDict(context, outlineRef);
  if (!root) {
    report.errors.push("Catalog /Outlines does not resolve to a dictionary.");
    return report;
  }
  const seen = new Set();
  walkOutlineList({ context, firstRef: root.get(PDFName.of("First")), parentRef: outlineRef, pageRefs, depth: 1, report, seen });
  const lastRef = root.get(PDFName.of("Last"));
  if (report.nodeCount > 0 && !(lastRef instanceof PDFRef)) report.errors.push("Outline root is missing a valid /Last reference.");
  const count = root.lookupMaybe(PDFName.of("Count"), PDFNumber)?.asNumber();
  if (count !== undefined && Math.abs(count) < report.nodeCount) report.errors.push(`Outline root /Count ${count} is smaller than discovered node count ${report.nodeCount}.`);
  return report;
}

function walkOutlineList({ context, firstRef, parentRef, pageRefs, depth, report, seen }) {
  let currentRef = firstRef;
  let previousRef;
  while (currentRef instanceof PDFRef) {
    const currentKey = refKey(currentRef);
    if (seen.has(currentKey)) {
      report.errors.push(`Outline cycle detected at ${currentKey}.`);
      return;
    }
    seen.add(currentKey);
    const node = lookupDict(context, currentRef);
    if (!node) {
      report.errors.push(`Outline node ${currentKey} does not resolve to a dictionary.`);
      return;
    }
    report.nodeCount += 1;
    report.maxDepth = Math.max(report.maxDepth, depth);
    const title = readPdfText(node.get(PDFName.of("Title")));
    report.titles.push(title || "(untitled)");
    if (refKey(node.get(PDFName.of("Parent"))) !== refKey(parentRef)) report.errors.push(`Outline node ${title || currentKey} has an invalid /Parent.`);
    if (previousRef && refKey(node.get(PDFName.of("Prev"))) !== refKey(previousRef)) report.errors.push(`Outline node ${title || currentKey} has an invalid /Prev.`);
    const dest = node.lookupMaybe(PDFName.of("Dest"), PDFArray);
    const destPageRef = dest?.get(0);
    if (!(destPageRef instanceof PDFRef) || !pageRefs.has(refKey(destPageRef))) report.errors.push(`Outline node ${title || currentKey} has an unresolved /Dest page.`);
    const firstChildRef = node.get(PDFName.of("First"));
    const lastChildRef = node.get(PDFName.of("Last"));
    if (firstChildRef || lastChildRef) {
      if (!(firstChildRef instanceof PDFRef) || !(lastChildRef instanceof PDFRef)) report.errors.push(`Outline node ${title || currentKey} has incomplete child references.`);
      walkOutlineList({ context, firstRef: firstChildRef, parentRef: currentRef, pageRefs, depth: depth + 1, report, seen });
    }
    previousRef = currentRef;
    currentRef = node.get(PDFName.of("Next"));
  }
}

function inspectLinkAnnotations(pdf, pages, pageRefs) {
  const report = { linkCount: 0, errors: [] };
  pages.forEach((page, pageIndex) => {
    const annots = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
    if (!annots) return;
    const size = page.getSize();
    for (let index = 0; index < annots.size(); index += 1) {
      const annotation = lookupDict(pdf.context, annots.get(index));
      if (!annotation || annotation.lookupMaybe(PDFName.of("Subtype"), PDFName)?.asString() !== "/Link") continue;
      report.linkCount += 1;
      const rect = annotation.lookupMaybe(PDFName.of("Rect"), PDFArray);
      if (!rect || rect.size() !== 4) {
        report.errors.push(`Link annotation ${index} on page ${pageIndex + 1} is missing a four-value /Rect.`);
      } else {
        const values = [0, 1, 2, 3].map((item) => rect.lookup(item, PDFNumber).asNumber());
        const [x1, y1, x2, y2] = values;
        if (x2 <= x1 || y2 <= y1) report.errors.push(`Link annotation ${index} on page ${pageIndex + 1} has a non-positive rectangle.`);
        if (x1 < -1 || y1 < -1 || x2 > size.width + 1 || y2 > size.height + 1) report.errors.push(`Link annotation ${index} on page ${pageIndex + 1} is outside the page bounds.`);
      }
      const dest = annotation.lookupMaybe(PDFName.of("Dest"), PDFArray);
      const destPageRef = dest?.get(0);
      if (!(destPageRef instanceof PDFRef) || !pageRefs.has(refKey(destPageRef))) report.errors.push(`Link annotation ${index} on page ${pageIndex + 1} has an unresolved destination.`);
    }
  });
  return report;
}

function inspectExpectedText(filePath, expectedText) {
  if (expectedText.length === 0) return { checked: false, errors: [] };
  const result = spawnSync("pdftotext", [filePath, "-"], { encoding: "utf8" });
  if (result.error?.code === "ENOENT") return { checked: false, errors: ["pdftotext is unavailable; cannot validate expected text."] };
  if (result.status !== 0) return { checked: false, errors: [result.stderr || "pdftotext failed."] };
  const errors = expectedText.filter((text) => !result.stdout.includes(text)).map((text) => `Expected text not found in exported PDF: ${text}`);
  return { checked: true, expectedText, errors };
}

function renderRepresentativePages(filePath, pageCount) {
  const renderDir = "/private/tmp/publish-pro-validation/navigation-renders";
  mkdirSync(renderDir, { recursive: true });
  const prefix = join(renderDir, basename(filePath, ".pdf"));
  const result = spawnSync("pdftoppm", ["-png", "-f", "1", "-l", String(Math.min(pageCount, 4)), filePath, prefix], { encoding: "utf8" });
  if (result.error?.code === "ENOENT") return { rendered: false, reason: "pdftoppm is unavailable." };
  if (result.status !== 0) return { rendered: false, reason: result.stderr || "pdftoppm failed." };
  return { rendered: true, prefix };
}

function lookupDict(context, object) {
  if (object instanceof PDFRef) return context.lookup(object, PDFDict);
  return object instanceof PDFDict ? object : undefined;
}

function readPdfText(value) {
  if (value instanceof PDFString || value instanceof PDFHexString) return value.decodeText();
  return "";
}

function refKey(ref) {
  return ref instanceof PDFRef ? `${ref.objectNumber} ${ref.generationNumber}` : "";
}

function fail(message) {
  console.error(message);
  process.exit(1);
}
