import { strFromU8, unzipSync } from "fflate";

const LARGE_DECK_BYTES = 100 * 1024 * 1024;
const MANY_SLIDE_THRESHOLD = 250;
const MANY_MEDIA_THRESHOLD = 500;
const SLIDE_RENDER_LIMIT = 250;
const TEXT_RENDER_LIMIT = 120;
const WARNING_RENDER_LIMIT = 80;
const ZIP_SIGNATURE_1 = 0x50;
const ZIP_SIGNATURE_2 = 0x4b;
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0];

export interface PptxSlide {
  number: number;
  path: string;
  title: string;
  texts: string[];
  renderedTexts: string[];
  textCount: number;
  notes: string[];
  renderedNotes: string[];
  noteCount: number;
  mediaRefs: string[];
  relationshipCount: number;
  externalRelationshipCount: number;
  warnings: string[];
}

export interface PptxMedia {
  path: string;
  name: string;
  extension: string;
  size: number;
  contentType: string;
}

export interface PptxPackageEntry {
  path: string;
  size: number;
  directory: boolean;
}

export interface PptxSummary {
  slideCount: number;
  renderedSlideCount: number;
  mediaCount: number;
  noteSlideCount: number;
  relationshipCount: number;
  externalRelationshipCount: number;
  packageEntryCount: number;
}

export interface ParsedPptx {
  title: string;
  app: string;
  slides: PptxSlide[];
  renderedSlides: PptxSlide[];
  media: PptxMedia[];
  packageEntries: PptxPackageEntry[];
  warnings: string[];
  summary: PptxSummary;
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode: string;
}

interface ContentTypes {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

type ZipEntries = Record<string, Uint8Array>;

export async function parsePptx(data: ArrayBuffer): Promise<ParsedPptx> {
  validatePptxSignature(data);

  const bytes = new Uint8Array(data);
  const zip = unzipSync(bytes);
  const fileNames = Object.keys(zip);
  const warnings: string[] = [];

  if (!zip["[Content_Types].xml"]) {
    throw new Error("Package is missing [Content_Types].xml.");
  }
  if (!zip["ppt/presentation.xml"]) {
    throw new Error("Package is missing ppt/presentation.xml.");
  }
  if (bytes.byteLength > LARGE_DECK_BYTES) {
    warnings.push(`Deck is ${formatBytes(bytes.byteLength)}; rendering is capped for responsiveness.`);
  }

  const contentTypes = parseContentTypes(await readText(zip, "[Content_Types].xml"));
  const docProps = await readOptionalText(zip, "docProps/core.xml");
  const appProps = await readOptionalText(zip, "docProps/app.xml");
  const presentationRels = await readRelationships(zip, "ppt/_rels/presentation.xml.rels");
  const slidePaths = collectSlidePaths(zip);
  const slides = await Promise.all(slidePaths.map((path, index) => parseSlide(zip, path, index + 1, contentTypes)));
  const media = collectMedia(zip, contentTypes);
  const packageEntries = fileNames
    .map((path) => ({
      directory: path.endsWith("/"),
      path,
      size: path.endsWith("/") ? 0 : zip[path]?.byteLength ?? 0,
    }))
    .sort((a, b) => a.path.localeCompare(b.path));

  const noteSlideCount = fileNames.filter((path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(path)).length;
  const relationshipFiles = await Promise.all(
    fileNames.filter((path) => path.endsWith(".rels")).map((path) => readRelationships(zip, path)),
  );
  const allRelationships = relationshipFiles.flat();
  const externalRelationshipCount = allRelationships.filter((rel) => rel.targetMode.toLowerCase() === "external").length;

  if (slidePaths.length === 0) warnings.push("Deck has no slide XML files.");
  if (slidePaths.length > MANY_SLIDE_THRESHOLD) warnings.push(`${slidePaths.length} slides found; rendered slide list is capped.`);
  if (media.length > MANY_MEDIA_THRESHOLD) warnings.push(`${media.length} media files found; media list is capped.`);
  if (externalRelationshipCount > 0) warnings.push(`${externalRelationshipCount} external relationships are listed as metadata only.`);
  if (fileNames.some((path) => /^ppt\/embeddings\//i.test(path))) {
    warnings.push("Embedded OLE/package content detected; embedded objects are not opened or executed.");
  }
  if (fileNames.some((path) => /vbaProject\.bin$/i.test(path))) {
    warnings.push("VBA macro project detected; macros are not opened or executed.");
  }

  const relationshipCount = allRelationships.length + presentationRels.length;
  const renderedSlides = slides.slice(0, SLIDE_RENDER_LIMIT);

  return {
    app: firstXmlText(appProps ?? "", "Application"),
    media,
    packageEntries,
    renderedSlides,
    slides,
    summary: {
      externalRelationshipCount,
      mediaCount: media.length,
      noteSlideCount,
      packageEntryCount: packageEntries.length,
      relationshipCount,
      renderedSlideCount: renderedSlides.length,
      slideCount: slides.length,
    },
    title: firstXmlText(docProps ?? "", "dc:title"),
    warnings: warnings.slice(0, WARNING_RENDER_LIMIT),
  };
}

export function filterSlides(slides: PptxSlide[], query: string): PptxSlide[] {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return slides;
  return slides.filter((slide) => {
    const haystack = [slide.title, ...slide.texts, ...slide.notes, ...slide.mediaRefs].join(" ").toLowerCase();
    return haystack.includes(normalized);
  });
}

export function filterMedia(media: PptxMedia[], query: string): PptxMedia[] {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return media;
  const extensionQuery = normalized.startsWith(".") ? normalized.slice(1) : "";
  return media.filter((item) => {
    return (
      item.path.toLowerCase().includes(normalized) ||
      item.contentType.toLowerCase().includes(normalized) ||
      Boolean(extensionQuery && item.extension.toLowerCase() === extensionQuery)
    );
  });
}

function validatePptxSignature(data: ArrayBuffer): void {
  if (data.byteLength < 4) {
    throw new Error("File is too small to be a .pptx package.");
  }
  const bytes = new Uint8Array(data, 0, 4);
  if (CFB_SIGNATURE.every((value, index) => bytes[index] === value)) {
    throw new Error("Encrypted or legacy PowerPoint files are not supported in v0.1.");
  }
  if (bytes[0] !== ZIP_SIGNATURE_1 || bytes[1] !== ZIP_SIGNATURE_2) {
    throw new Error("File is not a valid .pptx zip package.");
  }
}

async function parseSlide(zip: ZipEntries, path: string, number: number, contentTypes: ContentTypes): Promise<PptxSlide> {
  const xml = await readText(zip, path);
  const texts = extractTextRuns(xml);
  const relationships = await readRelationships(zip, slideRelationshipsPath(path));
  const notesPath = resolveNotesPath(path, relationships);
  const notes = notesPath ? extractTextRuns(await readOptionalText(zip, notesPath) ?? "") : [];
  const mediaRefs = relationships
    .filter((rel) => rel.type.toLowerCase().includes("/image") || rel.target.toLowerCase().includes("/media/"))
    .map((rel) => normalizePackagePath(path, rel.target))
    .sort((a, b) => a.localeCompare(b));
  const externalRelationshipCount = relationships.filter((rel) => rel.targetMode.toLowerCase() === "external").length;
  const warnings: string[] = [];

  if (texts.length === 0) warnings.push("Slide has no extracted text.");
  if (notesPath && notes.length === 0) warnings.push("Speaker notes file is present but contains no extracted text.");
  if (externalRelationshipCount > 0) warnings.push(`${externalRelationshipCount} external relationships are listed as metadata only.`);
  relationships
    .filter((rel) => rel.type.toLowerCase().includes("oleobject") || rel.type.toLowerCase().includes("package"))
    .forEach(() => warnings.push("Embedded object relationship detected; it is not opened or executed."));

  return {
    externalRelationshipCount,
    mediaRefs,
    noteCount: notes.length,
    notes,
    number,
    path,
    relationshipCount: relationships.length,
    renderedNotes: notes.slice(0, TEXT_RENDER_LIMIT),
    renderedTexts: texts.slice(0, TEXT_RENDER_LIMIT),
    textCount: texts.length,
    texts,
    title: inferSlideTitle(texts, path, contentTypes),
    warnings,
  };
}

function collectSlidePaths(zip: ZipEntries): string[] {
  return Object.keys(zip)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path))
    .sort((a, b) => slideNumber(a) - slideNumber(b));
}

function collectMedia(zip: ZipEntries, contentTypes: ContentTypes): PptxMedia[] {
  return Object.entries(zip)
    .filter(([path]) => /^ppt\/media\//i.test(path) && !path.endsWith("/"))
    .map(([path, content]) => {
      const extension = extensionForPath(path);
      return {
        contentType: contentTypeForPath(path, contentTypes),
        extension,
        name: path.split("/").pop() ?? path,
        path,
        size: content.byteLength,
      };
    })
    .sort((a, b) => a.path.localeCompare(b.path));
}

async function readText(zip: ZipEntries, path: string): Promise<string> {
  const file = zip[path];
  if (!file) throw new Error(`Package is missing ${path}.`);
  return strFromU8(file);
}

async function readOptionalText(zip: ZipEntries, path: string): Promise<string | null> {
  const file = zip[path];
  return file ? strFromU8(file) : null;
}

async function readRelationships(zip: ZipEntries, path: string): Promise<Relationship[]> {
  const xml = await readOptionalText(zip, path);
  if (!xml) return [];
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/gi)].map((match) => {
    const attrs = parseAttributes(match[1] ?? "");
    return {
      id: attrs.get("Id") ?? "",
      target: attrs.get("Target") ?? "",
      targetMode: attrs.get("TargetMode") ?? "",
      type: attrs.get("Type") ?? "",
    };
  });
}

function parseContentTypes(xml: string): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();

  for (const match of xml.matchAll(/<Default\b([^>]*)\/?>/gi)) {
    const attrs = parseAttributes(match[1] ?? "");
    const extension = attrs.get("Extension");
    const contentType = attrs.get("ContentType");
    if (extension && contentType) defaults.set(extension.toLowerCase(), contentType);
  }
  for (const match of xml.matchAll(/<Override\b([^>]*)\/?>/gi)) {
    const attrs = parseAttributes(match[1] ?? "");
    const partName = attrs.get("PartName");
    const contentType = attrs.get("ContentType");
    if (partName && contentType) overrides.set(partName.replace(/^\//, ""), contentType);
  }

  return { defaults, overrides };
}

function extractTextRuns(xml: string): string[] {
  return [...xml.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/gi)]
    .map((match) => decodeXmlEntities(match[1] ?? "").trim())
    .filter(Boolean);
}

function firstXmlText(xml: string, tagName: string): string {
  const escaped = tagName.replace(":", "\\:");
  const match = xml.match(new RegExp(`<${escaped}\\b[^>]*>([\\s\\S]*?)<\\/${escaped}>`, "i"));
  return match ? decodeXmlEntities(match[1] ?? "").trim() : "";
}

function parseAttributes(raw: string): Map<string, string> {
  const attrs = new Map<string, string>();
  for (const match of raw.matchAll(/([A-Za-z_:][\w:.-]*)\s*=\s*"([^"]*)"/g)) {
    attrs.set(match[1], decodeXmlEntities(match[2] ?? ""));
  }
  return attrs;
}

function inferSlideTitle(texts: string[], path: string, _contentTypes: ContentTypes): string {
  const first = texts.find((text) => text.trim().length > 0);
  if (first) return first.length > 80 ? `${first.slice(0, 77)}...` : first;
  return `Slide ${slideNumber(path)}`;
}

function resolveNotesPath(slidePath: string, relationships: Relationship[]): string {
  const rel = relationships.find((relationship) => relationship.type.toLowerCase().includes("/notesslide"));
  return rel ? normalizePackagePath(slidePath, rel.target) : "";
}

function slideRelationshipsPath(slidePath: string): string {
  const name = slidePath.split("/").pop() ?? slidePath;
  return slidePath.replace(/[^/]+$/, `_rels/${name}.rels`);
}

function normalizePackagePath(fromPath: string, target: string): string {
  if (!target) return "";
  if (target.startsWith("/")) return target.slice(1);
  const base = fromPath.split("/").slice(0, -1);
  const parts = target.split("/");
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") base.pop();
    else base.push(part);
  }
  return base.join("/");
}

function contentTypeForPath(path: string, contentTypes: ContentTypes): string {
  const override = contentTypes.overrides.get(path);
  if (override) return override;
  return contentTypes.defaults.get(extensionForPath(path).toLowerCase()) ?? "";
}

function extensionForPath(path: string): string {
  const name = path.split("/").pop() ?? path;
  const dot = name.lastIndexOf(".");
  if (dot <= 0 || dot === name.length - 1) return "";
  return name.slice(dot + 1);
}

function slideNumber(path: string): number {
  const match = path.match(/slide(\d+)\.xml$/i);
  return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
}

function decodeXmlEntities(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  const units = ["KB", "MB", "GB"];
  let value = bytes / 1024;
  let unit = units.shift() ?? "KB";
  while (value >= 1024 && units.length > 0) {
    value /= 1024;
    unit = units.shift() ?? unit;
  }
  return `${value.toFixed(value >= 10 ? 0 : 1)} ${unit}`;
}
