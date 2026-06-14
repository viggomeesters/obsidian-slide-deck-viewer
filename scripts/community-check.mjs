import fs from "node:fs";
import https from "node:https";

const manifest = JSON.parse(fs.readFileSync("manifest.json", "utf8"));
const packageJson = JSON.parse(fs.readFileSync("package.json", "utf8"));
const versions = JSON.parse(fs.readFileSync("versions.json", "utf8"));
const source = fs.readFileSync("src/main.ts", "utf8");

const expectedRepo = "https://github.com/viggomeesters/obsidian-slide-deck-viewer";
const allowedManifestNamePattern = /^[A-Za-z0-9 ()+-]+$/;
const lowerName = manifest.name.toLowerCase();
const lowerDescription = manifest.description.toLowerCase();

const checks = [
  [fs.existsSync("README.md"), "README.md exists at repository root"],
  [fs.existsSync("LICENSE"), "LICENSE exists at repository root"],
  [fs.existsSync("manifest.json"), "manifest.json exists at repository root"],
  [fs.existsSync("main.js"), "main.js exists at repository root"],
  [fs.existsSync("styles.css"), "styles.css exists at repository root"],
  [fs.existsSync("CHANGELOG.md"), "CHANGELOG.md exists at repository root"],
  [fs.existsSync("SECURITY.md"), "SECURITY.md exists at repository root"],
  [fs.existsSync("CONTRIBUTING.md"), "CONTRIBUTING.md exists at repository root"],
  [fs.existsSync("docs/community-submission.md"), "community submission notes exist"],
  [/^\d+\.\d+\.\d+$/.test(manifest.version), "manifest version uses x.y.z SemVer"],
  [manifest.version === packageJson.version, "manifest and package versions match"],
  [versions[manifest.version] === manifest.minAppVersion, "versions.json maps manifest version to minAppVersion"],
  [/^[a-z-]+$/.test(manifest.id), "manifest id contains only lowercase letters and hyphens"],
  [!manifest.id.includes("obsidian"), "manifest id does not contain obsidian"],
  [!manifest.id.endsWith("plugin"), "manifest id does not end with plugin"],
  [manifest.id === "slide-deck-viewer", "manifest id is slide-deck-viewer"],
  [manifest.name === "Slide Deck Viewer", "manifest name is Slide Deck Viewer"],
  [allowedManifestNamePattern.test(manifest.name), "manifest name uses allowed punctuation"],
  [!lowerName.includes("obsidian") && !lowerName.includes("obsi-") && !lowerName.includes("-sidian"), "manifest name avoids obsidian"],
  [typeof manifest.description === "string" && manifest.description.length > 0, "manifest description is present"],
  [!lowerDescription.includes("obsidian"), "manifest description avoids obsidian"],
  [typeof manifest.author === "string" && manifest.author.length > 0, "manifest author is present"],
  [typeof manifest.minAppVersion === "string" && manifest.minAppVersion.length > 0, "manifest minAppVersion is present"],
  [typeof manifest.isDesktopOnly === "boolean", "manifest isDesktopOnly is boolean"],
  [!source.includes("detachLeavesOfType"), "does not detach leaves in onunload"],
  [!source.includes("revealLeaf("), "does not use revealLeaf with current minAppVersion"],
  [fs.existsSync("assets/hero.svg"), "repo hero asset exists"],
  [fs.existsSync("assets/screenshot.svg"), "repo screenshot asset exists"],
  [fs.existsSync("assets/social-preview.svg"), "repo social preview asset exists"],
];

const failures = checks.filter(([passes]) => !passes).map(([, label]) => label);
failures.push(...await checkCommunityDirectory());

if (failures.length > 0) {
  for (const failure of failures) console.error(`FAIL: ${failure}`);
  process.exit(1);
}

console.log("Obsidian community submission checks passed.");

async function checkCommunityDirectory() {
  if (process.env.SKIP_LIVE_COMMUNITY_CHECK === "1") return [];

  const failures = [];
  const plugins = await fetchJson("https://raw.githubusercontent.com/obsidianmd/obsidian-releases/master/community-plugins.json");
  const rawIdMatch = plugins.find((plugin) => plugin.id === manifest.id);
  const rawNameMatch = plugins.find((plugin) => plugin.name?.toLowerCase() === lowerName);

  if (rawIdMatch && rawIdMatch.repo !== expectedRepo) {
    failures.push(`community directory already has id ${manifest.id}`);
  }
  if (rawNameMatch && rawNameMatch.repo !== expectedRepo) {
    failures.push(`community directory already has name ${manifest.name}`);
  }

  const html = await fetchText(`https://community.obsidian.md/plugins/${manifest.id}`);
  const isNotFound = html.includes("<title>Plugin not found</title>");
  if (!isNotFound && !html.includes(expectedRepo)) {
    failures.push(`live community slug ${manifest.id} is already taken by another plugin`);
  }

  return failures;
}

function fetchText(url) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: { "User-Agent": "Mozilla/5.0" } }, (response) => {
      let body = "";
      response.setEncoding("utf8");
      response.on("data", (chunk) => { body += chunk; });
      response.on("end", () => resolve(body));
    }).on("error", reject);
  });
}

async function fetchJson(url) {
  return JSON.parse(await fetchText(url));
}
