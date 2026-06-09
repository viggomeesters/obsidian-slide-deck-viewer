import fs from "node:fs";

const manifest = JSON.parse(fs.readFileSync("manifest.json", "utf8"));
const versions = JSON.parse(fs.readFileSync("versions.json", "utf8"));
const main = fs.readFileSync("src/main.ts", "utf8");
const parser = fs.readFileSync("src/parser.ts", "utf8");
const styles = fs.readFileSync("styles.css", "utf8");
const bundle = fs.readFileSync("main.js", "utf8");

const forbidden = [
  "fetch(",
  "XMLHttpRequest",
  "WebSocket",
  "navigator.clipboard",
  "child_process",
  "spawn(",
  "exec(",
  "eval(",
  "new Function",
];

const assertions = [
  [manifest.id === "pptx-viewer", "manifest id is pptx-viewer"],
  [manifest.name === "PPTX Viewer", "manifest name is PPTX Viewer"],
  [manifest.version === "0.1.0", "manifest version is 0.1.0"],
  [versions[manifest.version] === manifest.minAppVersion, "versions.json maps manifest version"],
  [!/obsidian/i.test(manifest.description), "manifest description avoids product name"],
  [/^[a-z-]+$/.test(manifest.id) && !manifest.id.includes("obsidian") && !manifest.id.endsWith("plugin"), "manifest id follows directory rules"],
  [main.includes('const PPTX_EXTENSIONS = ["pptx"]'), "pptx extension is the v0.1 scope"],
  [main.includes("registerExtensions(PPTX_EXTENSIONS"), "pptx extension is registered"],
  [main.includes("extends FileView"), "FileView is used"],
  [main.includes("this.app.vault.readBinary(file)"), "vault binary reader is used"],
  [parser.includes("unzipSync") && parser.includes("strFromU8"), "local zip parser dependency is used"],
  [parser.includes("Encrypted or legacy PowerPoint files are not supported"), "encrypted/legacy state exists"],
  [parser.includes("Embedded OLE/package content detected"), "embedded OLE warning exists"],
  [parser.includes("SLIDE_RENDER_LIMIT = 250"), "slide render cap exists"],
  [forbidden.every((token) => !main.includes(token) && !parser.includes(token) && !bundle.includes(token)), "forbidden runtime APIs are absent"],
  [!styles.includes("!important"), "styles do not use important overrides"],
  [fs.existsSync("fixtures/simple.pptx"), "simple pptx fixture exists"],
  [fs.existsSync("fixtures/no-notes.pptx"), "no-notes pptx fixture exists"],
  [fs.existsSync("fixtures/embedded-object.pptx"), "embedded object pptx fixture exists"],
  [fs.existsSync("fixtures/large.pptx"), "large pptx fixture exists"],
  [fs.existsSync("fixtures/malformed.pptx"), "malformed pptx fixture exists"],
  [fs.existsSync("fixtures/encrypted.pptx"), "encrypted pptx fixture exists"],
  [fs.existsSync("README.md") && fs.existsSync("SECURITY.md") && fs.existsSync("LICENSE"), "docs and license exist"],
  [fs.existsSync("assets/hero.svg") && fs.existsSync("assets/social-preview.svg") && fs.existsSync("assets/screenshot.svg"), "visual assets exist"],
  [fs.existsSync("main.js") && fs.existsSync("styles.css") && fs.existsSync("manifest.json"), "release assets exist"],
];

const failures = assertions.filter(([passes]) => !passes).map(([, label]) => label);

if (failures.length > 0) {
  for (const failure of failures) {
    console.error(`FAIL: ${failure}`);
  }
  process.exit(1);
}

console.log("PPTX Viewer smoke checks passed.");
