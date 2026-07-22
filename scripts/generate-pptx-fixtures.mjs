import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, zipSync } from "fflate";

const outputDir = process.argv[2] || "/private/tmp/publish-pro-pptx-fixtures";
mkdirSync(outputDir, { recursive: true });

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";

const fixtures = [
  { name: "basic-deck.pptx", slides: [titleSlide("Publish Pro PPTX Fixture", "Title slide"), textSlide("Text Slide", "Multi-line text with Unicode: Cafe resume EUR GBP alpha beta."), shapeSlide(), tableSlide()] },
  { name: "navigation-deck.pptx", slides: [titleSlide("Navigation Deck", "Includes speaker notes"), hiddenSlide("Hidden Slide"), textSlide("Linked Slide", "External link fixture")], notes: { 1: "Speaker notes for the navigation fixture." }, hidden: [2], hyperlink: true },
  { name: "complex-deck.pptx", slides: Array.from({ length: 12 }, (_, index) => textSlide(`Slide ${index + 1}`, `Generated content slide ${index + 1}.`)) },
];

for (const fixture of fixtures) {
  const path = join(outputDir, fixture.name);
  writeFileSync(path, createPptx(fixture));
  console.log(path);
}
writeFileSync(join(outputDir, "corrupt.pptx"), strToU8("not a zip"));
console.log(join(outputDir, "corrupt.pptx"));

function createPptx(fixture) {
  const files = {
    "[Content_Types].xml": contentTypes(fixture.slides.length, fixture.notes || {}),
    "_rels/.rels": xml(`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`),
    "docProps/core.xml": xml(`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${escapeXml(fixture.name.replace(".pptx", ""))}</dc:title><dc:creator>Publish Pro</dc:creator></cp:coreProperties>`),
    "ppt/presentation.xml": presentationXml(fixture),
    "ppt/_rels/presentation.xml.rels": presentationRels(fixture.slides.length),
    "ppt/slideMasters/slideMaster1.xml": xml(`<p:sldMaster xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree/></p:cSld></p:sldMaster>`),
    "ppt/slideLayouts/slideLayout1.xml": xml(`<p:sldLayout xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree/></p:cSld></p:sldLayout>`),
    "ppt/theme/theme1.xml": xml(`<a:theme xmlns:a="${A}" name="Publish Pro Fixture"><a:themeElements/></a:theme>`),
  };
  fixture.slides.forEach((slide, index) => {
    const number = index + 1;
    files[`ppt/slides/slide${number}.xml`] = slideXml(slide, fixture.hidden?.includes(number));
    files[`ppt/slides/_rels/slide${number}.xml.rels`] = slideRels(number, fixture);
    if (fixture.notes?.[number]) files[`ppt/notesSlides/notesSlide${number}.xml`] = notesXml(fixture.notes[number]);
  });
  return zipSync(Object.fromEntries(Object.entries(files).map(([path, value]) => [path, typeof value === "string" ? strToU8(value) : value])));
}

function contentTypes(slideCount, notes) {
  const slides = Array.from({ length: slideCount }, (_, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("");
  const notesParts = Object.keys(notes).map((number) => `<Override PartName="/ppt/notesSlides/notesSlide${number}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`).join("");
  return xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${slides}${notesParts}</Types>`);
}

function presentationXml(fixture) {
  const slideIds = fixture.slides.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`).join("");
  return xml(`<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldMasterIdLst><p:sldMasterId id="1" r:id="rIdMaster1"/></p:sldMasterIdLst><p:sldIdLst>${slideIds}</p:sldIdLst><p:extLst><p:ext><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:section name="Fixture Section" id="{11111111-1111-1111-1111-111111111111}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/></p:presentation>`);
}

function presentationRels(slideCount) {
  const slides = Array.from({ length: slideCount }, (_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`).join("");
  return xml(`<Relationships xmlns="${REL}">${slides}<Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>`);
}

function slideRels(number, fixture) {
  const notes = fixture.notes?.[number] ? `<Relationship Id="rIdNotes${number}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${number}.xml"/>` : "";
  const link = fixture.hyperlink && number === 3 ? `<Relationship Id="rIdHyper1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/publish-pro-pptx" TargetMode="External"/>` : "";
  return xml(`<Relationships xmlns="${REL}"><Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>${notes}${link}</Relationships>`);
}

function slideXml(body, hidden) {
  return xml(`<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"${hidden ? ' show="0"' : ""}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg><p:spTree>${body}</p:spTree></p:cSld></p:sld>`);
}

function titleSlide(title, subtitle) {
  return textBox(title, 700000, 700000, 9000000, 900000, 3200) + textBox(subtitle, 800000, 1700000, 8400000, 700000, 1800);
}

function textSlide(title, text) {
  return textBox(title, 600000, 500000, 9800000, 650000, 2600) + textBox(text, 800000, 1500000, 9000000, 2600000, 1500);
}

function hiddenSlide(title) {
  return textBox(title, 600000, 600000, 9000000, 900000, 2600);
}

function shapeSlide() {
  return textBox("Shape Slide", 600000, 500000, 6000000, 650000, 2600) + shape("rect", 1000000, 1800000, 2500000, 1200000, "FEE2E2") + shape("ellipse", 4200000, 1800000, 2200000, 1300000, "DBEAFE");
}

function tableSlide() {
  return textBox("Table Slide", 600000, 500000, 6000000, 650000, 2600) + `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Table"/></p:nvGraphicFramePr><p:xfrm><a:off x="900000" y="1600000"/><a:ext cx="7000000" cy="2200000"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>Header A</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>Header B</a:t></a:r></a:p></a:txBody></a:tc></a:tr><a:tr><a:tc><a:txBody><a:p><a:r><a:t>One</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>Two</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function textBox(text, x, y, cx, cy, size) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${Math.floor(Math.random() * 100000)}" name="Text Box"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="${size}"><a:solidFill><a:srgbClr val="111827"/></a:solidFill></a:rPr><a:t>${escapeXml(text)}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function shape(prst, x, y, cx, cy, fill) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${Math.floor(Math.random() * 100000)}" name="${prst}"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${prst}"/><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></p:spPr></p:sp>`;
}

function notesXml(text) {
  return xml(`<p:notes xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>${textBox(text, 0, 0, 6000000, 1000000, 1200)}</p:spTree></p:cSld></p:notes>`);
}

function xml(value) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${value}`;
}

function escapeXml(value) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
