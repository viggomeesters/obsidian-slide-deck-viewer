import fs from "node:fs";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const root = new URL("../fixtures/", import.meta.url);
fs.mkdirSync(root, { recursive: true });

function contentTypes(extraOverrides = "") {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
  ${extraOverrides}
</Types>`;
}

function rels(items) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${items.join("\n")}
</Relationships>`;
}

function slide(texts) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>${texts.map((text) => `<p:sp><p:txBody><a:p><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p></p:txBody></p:sp>`).join("")}</p:spTree></p:cSld>
</p:sld>`;
}

function notes(texts) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>${texts.map((text) => `<p:sp><p:txBody><a:p><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p></p:txBody></p:sp>`).join("")}</p:spTree></p:cSld>
</p:notes>`;
}

async function writeZip(name, build) {
  const zip = new JSZip();
  await build(zip);
  const buffer = await zip.generateAsync({ compression: "DEFLATE", type: "nodebuffer" });
  fs.writeFileSync(new URL(name, root), buffer);
}

async function addBasicDeck(zip) {
  zip.file("[Content_Types].xml", contentTypes());
  zip.file("_rels/.rels", rels([
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>',
  ]));
  zip.file("docProps/core.xml", '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Quarterly Review</dc:title></cp:coreProperties>');
  zip.file("docProps/app.xml", "<Properties><Application>Fixture Generator</Application></Properties>");
  zip.file("ppt/presentation.xml", "<p:presentation/>");
  zip.file("ppt/_rels/presentation.xml.rels", rels([
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>',
  ]));
  zip.file("ppt/slides/slide1.xml", slide(["Quarterly Review", "Revenue up", "Costs stable"]));
  zip.file("ppt/slides/slide2.xml", slide(["Roadmap", "Ship viewer", "Prepare release"]));
  zip.file("ppt/slides/_rels/slide1.xml.rels", rels([
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
  ]));
  zip.file("ppt/slides/_rels/slide2.xml.rels", rels([
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid" TargetMode="External"/>',
  ]));
  zip.file("ppt/notesSlides/notesSlide1.xml", notes(["Mention read-only scope", "No external viewer opens"]));
  zip.file("ppt/media/image1.png", Buffer.from("89504e470d0a1a0a0000000d49484452", "hex"));
}

await writeZip("simple.pptx", addBasicDeck);
await writeZip("no-notes.pptx", async (zip) => {
  await addBasicDeck(zip);
  zip.remove("ppt/notesSlides/notesSlide1.xml");
  zip.file("ppt/slides/_rels/slide1.xml.rels", rels([]));
});
await writeZip("embedded-object.pptx", async (zip) => {
  await addBasicDeck(zip);
  zip.file("ppt/embeddings/oleObject1.bin", "binary metadata only");
  zip.file("ppt/slides/_rels/slide2.xml.rels", rels([
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>',
  ]));
});
await writeZip("large.pptx", async (zip) => {
  zip.file("[Content_Types].xml", contentTypes());
  zip.file("ppt/presentation.xml", "<p:presentation/>");
  for (let index = 1; index <= 275; index += 1) {
    zip.file(`ppt/slides/slide${index}.xml`, slide([`Slide ${index}`, "Large deck fixture"]));
  }
});

fs.writeFileSync(new URL("malformed.pptx", root), Buffer.from("not a zip"));
fs.writeFileSync(new URL("encrypted.pptx", root), Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));

function escapeXml(value) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

console.log(`PPTX fixtures written to ${fileURLToPath(root)}`);
