import assert from "node:assert/strict";
import fs from "node:fs";
import esbuild from "esbuild";

await esbuild.build({
  bundle: true,
  entryPoints: ["src/parser.ts"],
  format: "esm",
  outfile: ".tmp-parser-test.mjs",
  platform: "node",
  target: "node20",
});

const {
  filterMedia,
  filterSlides,
  parsePptx,
} = await import(new URL("../.tmp-parser-test.mjs", import.meta.url));

function fixture(name) {
  const data = fs.readFileSync(new URL(`../fixtures/${name}`, import.meta.url));
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
}

const simple = await parsePptx(fixture("simple.pptx"));
assert.equal(simple.summary.slideCount, 2);
assert.equal(simple.summary.mediaCount, 1);
assert.equal(simple.summary.noteSlideCount, 1);
assert.equal(simple.slides[0].title, "Quarterly Review");
assert.ok(simple.slides[0].notes.includes("Mention read-only scope"));
assert.ok(simple.slides[0].mediaRefs.includes("ppt/media/image1.png"));
assert.equal(filterSlides(simple.slides, "roadmap").length, 1);
assert.equal(filterMedia(simple.media, ".png").length, 1);
assert.ok(simple.warnings.some((warning) => warning.includes("external relationships")));

const noNotes = await parsePptx(fixture("no-notes.pptx"));
assert.equal(noNotes.summary.noteSlideCount, 0);
assert.equal(noNotes.slides[0].noteCount, 0);

const embedded = await parsePptx(fixture("embedded-object.pptx"));
assert.ok(embedded.warnings.some((warning) => warning.includes("Embedded OLE")));
assert.ok(embedded.slides[1].warnings.some((warning) => warning.includes("Embedded object")));

const large = await parsePptx(fixture("large.pptx"));
assert.equal(large.summary.slideCount, 275);
assert.equal(large.summary.renderedSlideCount, 250);
assert.ok(large.warnings.some((warning) => warning.includes("rendered slide list is capped")));

await assert.rejects(() => parsePptx(fixture("malformed.pptx")), /valid .pptx zip package/i);
await assert.rejects(() => parsePptx(fixture("encrypted.pptx")), /Encrypted or legacy PowerPoint/i);

fs.rmSync(new URL("../.tmp-parser-test.mjs", import.meta.url));
console.log("PPTX parser fixture tests passed.");

