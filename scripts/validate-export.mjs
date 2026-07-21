#!/usr/bin/env node
import { existsSync, mkdirSync, statSync } from "node:fs";
import { basename, join } from "node:path";
import { spawnSync } from "node:child_process";
import { PDFDocument } from "pdf-lib";

const [, , filePath, expectedPageCount] = process.argv;

if (!filePath) {
  console.error("Usage: node scripts/validate-export.mjs <pdf-path> [expected-page-count]");
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
