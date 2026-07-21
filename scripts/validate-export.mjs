#!/usr/bin/env node
import { existsSync, mkdirSync, statSync, writeFileSync } from "node:fs";
import { basename, join } from "node:path";
import { spawnSync } from "node:child_process";
import { PDFDocument, StandardFonts, rgb, degrees } from "pdf-lib";

const [, , filePath, expectedPageCount] = process.argv;

if (filePath === "--generate-fixtures") {
  const fixtureDir = "/private/tmp/publish-pro-validation";
  mkdirSync(fixtureDir, { recursive: true });
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const portrait = pdf.addPage([612, 792]);
  portrait.drawText("Publish Pro portrait fixture", { x: 72, y: 720, size: 18, font, color: rgb(0, 0, 0) });
  portrait.drawText("Highlight, underline, strikethrough, comments, shapes.", { x: 72, y: 680, size: 12, font });
  const landscape = pdf.addPage([841.89, 595.28]);
  landscape.drawText("Publish Pro landscape fixture", { x: 72, y: 520, size: 18, font, color: rgb(0, 0, 0) });
  landscape.drawText("Mixed-size page validation.", { x: 72, y: 480, size: 12, font });
  const rotated = pdf.addPage([612, 792]);
  rotated.setRotation(degrees(90));
  rotated.drawText("Rotated fixture page", { x: 72, y: 720, size: 18, font, color: rgb(0, 0, 0) });
  const output = join(fixtureDir, "publish-pro-generated-fixture.pdf");
  writeFileSync(output, await pdf.save());
  console.log(output);
  process.exit(0);
}

if (!filePath) {
  console.error("Usage: node scripts/validate-export.mjs <pdf-path> [expected-page-count]\n       node scripts/validate-export.mjs --generate-fixtures");
  process.exit(1);
}

if (!existsSync(filePath)) {
  console.error(`PDF does not exist: ${filePath}`);
  process.exit(1);
}

const stats = statSync(filePath);
if (stats.size <= 0) {
  console.error(`PDF is empty: ${filePath}`);
  process.exit(1);
}

const bytes = await import("node:fs").then((fs) => fs.readFileSync(filePath));
const pdf = await PDFDocument.load(bytes);
const pageCount = pdf.getPageCount();

if (expectedPageCount !== undefined && pageCount !== Number(expectedPageCount)) {
  console.error(`Expected ${expectedPageCount} pages, found ${pageCount}.`);
  process.exit(1);
}

const dimensions = pdf.getPages().map((page, index) => {
  const size = page.getSize();
  return {
    page: index + 1,
    width: Number(size.width.toFixed(2)),
    height: Number(size.height.toFixed(2)),
    rotation: page.getRotation().angle,
    orientation: size.width >= size.height ? "landscape" : "portrait",
  };
});

const renderDir = "/private/tmp/publish-pro-validation/renders";
mkdirSync(renderDir, { recursive: true });
const renderPrefix = join(renderDir, basename(filePath, ".pdf"));
const render = spawnSync("pdftoppm", ["-png", "-f", "1", "-l", String(Math.min(pageCount, 2)), filePath, renderPrefix], {
  encoding: "utf8",
});

const rendered = render.status === 0;
if (!rendered && render.error?.code !== "ENOENT") {
  console.warn(render.stderr || render.error?.message);
}

console.log(
  JSON.stringify(
    {
      filePath,
      size: stats.size,
      pageCount,
      dimensions,
      rendered,
      renderPrefix: rendered ? renderPrefix : null,
    },
    null,
    2
  )
);
