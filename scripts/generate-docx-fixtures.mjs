import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, zipSync } from "fflate";

const outputDir = process.argv[2] || "/private/tmp/publish-pro-docx-fixtures";
mkdirSync(outputDir, { recursive: true });

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/package/2006/relationships";
const OR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const fixtures = [
  {
    name: "typography.docx",
    body: [
      heading("Typography Fixture", 1),
      paragraph("Helvetica-style paragraph with direct bold, italic, underline and colour."),
      paragraph("Unicode text: Cafe, resume, EUR, GBP, alpha beta, and non-breaking spaces."),
      paragraph("Superscript and subscript samples are represented as formatted runs in source XML."),
    ].join(""),
  },
  {
    name: "lists.docx",
    body: [
      heading("Lists Fixture", 1),
      list("First decimal item", "2", 0),
      list("Nested roman item", "3", 1),
      list("Nested alphabetic item", "4", 2),
      list("Second decimal item", "2", 0),
    ].join(""),
  },
  {
    name: "tables.docx",
    body: [
      heading("Tables Fixture", 1),
      table([
        ["Header A", "Header B", "Header C"],
        ["Merged-cell source is simplified", "Shaded cell", "Long table cell text that wraps across multiple lines inside Publish Pro."],
        ["Row 3", "Row 3B", "Row 3C"],
      ]),
    ].join(""),
  },
  {
    name: "sections.docx",
    body: [
      heading("Sections Fixture", 1),
      paragraph("Portrait section with first page header and footer relationships."),
      pageBreak(),
      heading("Landscape Section", 1),
      paragraph("This fixture includes a landscape section size in the section properties."),
    ].join(""),
    landscape: true,
  },
  {
    name: "review.docx",
    extraFiles: {
      "word/comments.xml": xml(`<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Publish Pro"><w:p><w:r><w:t>Review comment fixture</w:t></w:r></w:p></w:comment></w:comments>`),
      "word/footnotes.xml": xml(`<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:r><w:t>Footnote fixture text</w:t></w:r></w:p></w:footnote></w:footnotes>`),
    },
    body: [
      heading("Review Fixture", 1),
      paragraph("Paragraph with footnote reference."),
      `<w:p><w:ins><w:r><w:t>Inserted text</w:t></w:r></w:ins><w:del><w:r><w:t>Deleted text</w:t></w:r></w:del></w:p>`,
    ].join(""),
  },
  {
    name: "media-links.docx",
    relationships: `<Relationship Id="rIdHyper1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/publish-pro" TargetMode="External"/>`,
    body: [
      heading("Media and Links Fixture", 1),
      `<w:p><w:hyperlink r:id="rIdHyper1"><w:r><w:rPr><w:u w:val="single"/><w:color w:val="2563EB"/></w:rPr><w:t>Publish Pro link</w:t></w:r></w:hyperlink></w:p>`,
      paragraph("Unsupported EMF/WMF media references are represented by warnings when present in real DOCX packages."),
    ].join(""),
  },
];

for (const fixture of fixtures) {
  writeFileSync(join(outputDir, fixture.name), createDocx(fixture));
  console.log(join(outputDir, fixture.name));
}

function createDocx(fixture) {
  const files = {
    "[Content_Types].xml": contentTypes(fixture.extraFiles),
    "_rels/.rels": xml(`<Relationships xmlns="${R}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": xml(`<Relationships xmlns="${R}"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>${fixture.relationships || ""}</Relationships>`),
    "word/styles.xml": styles(),
    "word/numbering.xml": numbering(),
    "word/header1.xml": xml(`<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Fixture Header</w:t></w:r></w:p></w:hdr>`),
    "word/footer1.xml": xml(`<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Fixture Footer</w:t></w:r></w:p></w:ftr>`),
    "word/document.xml": documentXml(fixture),
    ...(fixture.extraFiles || {}),
  };
  return zipSync(Object.fromEntries(Object.entries(files).map(([path, value]) => [path, strToU8(value)])));
}

function contentTypes(extraFiles) {
  const extras = extraFiles ? Object.keys(extraFiles).map((path) => `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${path.includes("comment") ? "comments" : "footnotes"}+xml"/>`).join("") : "";
  return xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>${extras}</Types>`);
}

function documentXml(fixture) {
  const pgSz = fixture.landscape ? '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>' : '<w:pgSz w:w="12240" w:h="15840"/>';
  return xml(`<w:document xmlns:w="${W}" xmlns:r="${OR}"><w:body>${fixture.body}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>${pgSz}<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
}

function styles() {
  return xml(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/><w:keepNext/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style></w:styles>`);
}

function numbering() {
  return xml(`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%2."/></w:lvl><w:lvl w:ilvl="2"><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%3."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="4"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
}

function heading(text, level) {
  return `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function paragraph(text) {
  return `<w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function list(text, numId, level) {
  return `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function table(rows) {
  return `<w:tbl><w:tblGrid>${rows[0].map(() => '<w:gridCol w:w="2880"/>').join("")}</w:tblGrid>${rows.map((row, rowIndex) => `<w:tr>${row.map((cell) => `<w:tc><w:tcPr>${rowIndex === 1 && cell.includes("Shaded") ? '<w:shd w:fill="E2E8F0"/>' : ""}</w:tcPr>${paragraph(cell)}</w:tc>`).join("")}</w:tr>`).join("")}</w:tbl>`;
}

function pageBreak() {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

function xml(value) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${value}`;
}

function escapeXml(value) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
