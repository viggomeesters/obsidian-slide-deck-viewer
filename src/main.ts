import {
  FileView,
  Notice,
  Plugin,
  TFile,
  WorkspaceLeaf,
  setIcon,
} from "obsidian";
import { ParsedPptx, PptxMedia, PptxPackageEntry, PptxSlide, filterMedia, filterSlides, parsePptx } from "./parser";

const VIEW_TYPE_PPTX_VIEWER = "slide-deck-viewer";
const PPTX_EXTENSIONS = ["pptx"];
const MEDIA_RENDER_LIMIT = 120;
const PACKAGE_RENDER_LIMIT = 240;

export default class PptxViewerPlugin extends Plugin {
  async onload(): Promise<void> {
    this.registerView(
      VIEW_TYPE_PPTX_VIEWER,
      (leaf) => new PptxViewerView(leaf),
    );
    this.registerExtensions(PPTX_EXTENSIONS, VIEW_TYPE_PPTX_VIEWER);

    this.addCommand({
      id: "open-current-pptx-in-viewer",
      name: "Open current PPTX file in viewer",
      checkCallback: (checking) => {
        const file = this.app.workspace.getActiveFile();
        if (!isPptxFile(file)) return false;

        if (!checking) {
          void this.openPptxFile(file);
        }
        return true;
      },
    });
  }

  async openPptxFile(file: TFile): Promise<void> {
    const leaf = this.app.workspace.getLeaf(false);
    await leaf.setViewState({
      active: true,
      state: { file: file.path },
      type: VIEW_TYPE_PPTX_VIEWER,
    });
  }
}

class PptxViewerView extends FileView {
  private deck: ParsedPptx | null = null;
  private activeSlideNumber = 1;
  private filterValue = "";
  private errorMessage = "";

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
  }

  getViewType(): string {
    return VIEW_TYPE_PPTX_VIEWER;
  }

  getDisplayText(): string {
    return this.file?.basename ?? "slide deck viewer";
  }

  getIcon(): string {
    return "presentation";
  }

  async onLoadFile(file: TFile): Promise<void> {
    await this.loadDeck(file);
  }

  async onUnloadFile(): Promise<void> {
    this.deck = null;
    this.activeSlideNumber = 1;
    this.errorMessage = "";
    this.contentEl.empty();
  }

  private async loadDeck(file: TFile): Promise<void> {
    try {
      const data = await this.app.vault.readBinary(file);
      this.deck = await parsePptx(data);
      this.activeSlideNumber = this.deck.renderedSlides[0]?.number ?? 1;
      this.errorMessage = "";
    } catch (error) {
      this.deck = null;
      this.activeSlideNumber = 1;
      this.errorMessage = `Unable to read presentation: ${getErrorMessage(error)}`;
    }
    this.render();
  }

  private render(): void {
    const container = this.contentEl;
    container.empty();
    container.addClass("slide-deck-viewer");

    const header = container.createDiv({ cls: "slide-deck-viewer__header" });
    this.renderTitle(header);
    this.renderToolbar(header);

    if (!this.file) {
      renderMessage(container, "No PPTX file is attached to this viewer.");
      return;
    }
    if (!isPptxFile(this.file)) {
      renderMessage(container, "This viewer only supports .pptx files.");
      return;
    }
    if (this.errorMessage) {
      renderMessage(container, this.errorMessage);
      return;
    }
    if (!this.deck) {
      renderMessage(container, "Presentation is not loaded.");
      return;
    }

    renderSummary(container, this.deck);
    renderWarnings(container, this.deck.warnings, "Deck warnings");

    const body = container.createDiv({ cls: "slide-deck-viewer__body" });
    this.renderSlideList(body, this.deck);
    this.renderDetail(body, this.deck);
  }

  private renderTitle(parent: HTMLElement): void {
    const title = parent.createDiv({ cls: "slide-deck-viewer__title" });
    title.createDiv({
      cls: "slide-deck-viewer__filename",
      text: this.file?.name ?? "PPTX file",
    });
    title.createDiv({
      cls: "slide-deck-viewer__path",
      text: this.file?.path ?? "",
    });
  }

  private renderToolbar(parent: HTMLElement): void {
    const toolbar = parent.createDiv({ cls: "slide-deck-viewer__toolbar" });
    const searchWrap = toolbar.createDiv({ cls: "slide-deck-viewer__search" });
    setIcon(searchWrap.createSpan({ cls: "slide-deck-viewer__search-icon" }), "search");
    const searchInput = searchWrap.createEl("input", {
      attr: {
        "aria-label": "Filter slides and media",
        placeholder: "Filter",
        spellcheck: "false",
        type: "search",
        value: this.filterValue,
      },
    });
    searchInput.addEventListener("input", () => {
      this.filterValue = searchInput.value;
      this.render();
    });

    const refreshButton = createIconButton(toolbar, "refresh-cw", "Refresh presentation");
    refreshButton.addEventListener("click", () => {
      void this.reloadFile();
    });
  }

  private renderSlideList(parent: HTMLElement, deck: ParsedPptx): void {
    const sidebar = parent.createDiv({ cls: "slide-deck-viewer__sidebar" });
    sidebar.createDiv({ cls: "slide-deck-viewer__section-title", text: "Slides" });

    const slides = filterSlides(deck.renderedSlides, this.filterValue);
    if (slides.length === 0) {
      sidebar.createDiv({ cls: "slide-deck-viewer__empty", text: "No slides match the filter." });
      return;
    }

    slides.forEach((slide) => {
      const button = sidebar.createEl("button", {
        cls: "slide-deck-viewer__slide-button",
        attr: { type: "button" },
      });
      button.toggleClass("is-active", slide.number === this.activeSlideNumber);
      button.createSpan({ cls: "slide-deck-viewer__slide-number", text: String(slide.number) });
      const label = button.createSpan({ cls: "slide-deck-viewer__slide-label" });
      label.createSpan({ cls: "slide-deck-viewer__slide-title", text: slide.title });
      label.createSpan({
        cls: "slide-deck-viewer__slide-meta",
        text: `${slide.textCount} text runs, ${slide.noteCount} notes`,
      });
      button.addEventListener("click", () => {
        this.activeSlideNumber = slide.number;
        this.render();
      });
    });
  }

  private renderDetail(parent: HTMLElement, deck: ParsedPptx): void {
    const detail = parent.createDiv({ cls: "slide-deck-viewer__detail" });
    const activeSlide = deck.slides.find((slide) => slide.number === this.activeSlideNumber) ?? deck.renderedSlides[0];
    if (!activeSlide) {
      renderMessage(detail, "No slides are available.");
      return;
    }

    renderSlide(detail, activeSlide);
    renderMedia(detail, filterMedia(deck.media, this.filterValue));
    renderPackageEntries(detail, deck.packageEntries);
  }

  private async reloadFile(): Promise<void> {
    if (!this.file) {
      new Notice("No PPTX file to refresh");
      return;
    }
    await this.loadDeck(this.file);
  }
}

function renderSummary(parent: HTMLElement, deck: ParsedPptx): void {
  const summary = parent.createDiv({ cls: "slide-deck-viewer__summary" });
  summary.createSpan({ cls: "slide-deck-viewer__pill", text: `${deck.summary.slideCount} slides` });
  summary.createSpan({ cls: "slide-deck-viewer__pill", text: `${deck.summary.mediaCount} media` });
  summary.createSpan({ cls: "slide-deck-viewer__pill", text: `${deck.summary.noteSlideCount} note slides` });
  summary.createSpan({ cls: "slide-deck-viewer__pill", text: `${deck.summary.packageEntryCount} package entries` });
  if (deck.summary.externalRelationshipCount > 0) {
    summary.createSpan({ cls: "slide-deck-viewer__note", text: `${deck.summary.externalRelationshipCount} external relationships listed` });
  }
  if (deck.summary.renderedSlideCount < deck.summary.slideCount) {
    summary.createSpan({ cls: "slide-deck-viewer__note", text: `${deck.summary.renderedSlideCount} slides rendered` });
  }
}

function renderSlide(parent: HTMLElement, slide: PptxSlide): void {
  const section = parent.createDiv({ cls: "slide-deck-viewer__slide-detail" });
  const heading = section.createDiv({ cls: "slide-deck-viewer__detail-heading" });
  heading.createDiv({ cls: "slide-deck-viewer__detail-title", text: `Slide ${slide.number}: ${slide.title}` });
  heading.createDiv({ cls: "slide-deck-viewer__detail-path", text: slide.path });
  renderWarnings(section, slide.warnings, "Slide warnings");

  const textGrid = section.createDiv({ cls: "slide-deck-viewer__text-grid" });
  renderTextBlock(textGrid, "Text", slide.renderedTexts, slide.textCount);
  renderTextBlock(textGrid, "Speaker notes", slide.renderedNotes, slide.noteCount);

  if (slide.mediaRefs.length > 0) {
    const refs = section.createDiv({ cls: "slide-deck-viewer__refs" });
    refs.createDiv({ cls: "slide-deck-viewer__section-title", text: "Media references" });
    slide.mediaRefs.forEach((ref) => refs.createDiv({ cls: "slide-deck-viewer__ref", text: ref }));
  }
}

function renderTextBlock(parent: HTMLElement, title: string, lines: string[], totalCount: number): void {
  const block = parent.createDiv({ cls: "slide-deck-viewer__text-block" });
  block.createDiv({ cls: "slide-deck-viewer__section-title", text: title });
  if (lines.length === 0) {
    block.createDiv({ cls: "slide-deck-viewer__empty", text: "No extracted text." });
    return;
  }
  const list = block.createEl("ol", { cls: "slide-deck-viewer__text-list" });
  lines.forEach((line) => list.createEl("li", { text: line }));
  if (lines.length < totalCount) {
    block.createDiv({ cls: "slide-deck-viewer__note", text: `${totalCount - lines.length} additional text runs hidden` });
  }
}

function renderMedia(parent: HTMLElement, media: PptxMedia[]): void {
  const section = parent.createDiv({ cls: "slide-deck-viewer__panel" });
  section.createDiv({ cls: "slide-deck-viewer__section-title", text: "Media" });
  if (media.length === 0) {
    section.createDiv({ cls: "slide-deck-viewer__empty", text: "No media matches the current filter." });
    return;
  }
  const table = section.createEl("table", { cls: "slide-deck-viewer__table" });
  const head = table.createEl("thead").createEl("tr");
  ["Name", "Type", "Size"].forEach((label) => head.createEl("th", { text: label }));
  const body = table.createEl("tbody");
  media.slice(0, MEDIA_RENDER_LIMIT).forEach((item) => {
    const row = body.createEl("tr");
    row.createEl("td", { text: item.path });
    row.createEl("td", { text: item.contentType || item.extension || "unknown" });
    row.createEl("td", { text: formatBytes(item.size) });
  });
  if (media.length > MEDIA_RENDER_LIMIT) {
    section.createDiv({ cls: "slide-deck-viewer__note", text: `${media.length - MEDIA_RENDER_LIMIT} additional media files hidden` });
  }
}

function renderPackageEntries(parent: HTMLElement, entries: PptxPackageEntry[]): void {
  const section = parent.createDiv({ cls: "slide-deck-viewer__panel" });
  section.createDiv({ cls: "slide-deck-viewer__section-title", text: "Package diagnostics" });
  const table = section.createEl("table", { cls: "slide-deck-viewer__table" });
  const head = table.createEl("thead").createEl("tr");
  ["Path", "Kind", "Size"].forEach((label) => head.createEl("th", { text: label }));
  const body = table.createEl("tbody");
  entries.slice(0, PACKAGE_RENDER_LIMIT).forEach((entry) => {
    const row = body.createEl("tr");
    row.createEl("td", { text: entry.path });
    row.createEl("td", { text: entry.directory ? "directory" : "file" });
    row.createEl("td", { text: entry.directory ? "" : formatBytes(entry.size) });
  });
  if (entries.length > PACKAGE_RENDER_LIMIT) {
    section.createDiv({ cls: "slide-deck-viewer__note", text: `${entries.length - PACKAGE_RENDER_LIMIT} additional package entries hidden` });
  }
}

function renderWarnings(parent: HTMLElement, warnings: string[], title: string): void {
  if (warnings.length === 0) return;
  const box = parent.createDiv({ cls: "slide-deck-viewer__warnings" });
  box.createDiv({ cls: "slide-deck-viewer__warnings-title", text: title });
  warnings.slice(0, 8).forEach((warning) => box.createDiv({ cls: "slide-deck-viewer__warning", text: warning }));
  if (warnings.length > 8) {
    box.createDiv({ cls: "slide-deck-viewer__warning-more", text: `${warnings.length - 8} additional warnings hidden` });
  }
}

function createIconButton(parent: HTMLElement, icon: string, label: string): HTMLButtonElement {
  const button = parent.createEl("button", {
    attr: { "aria-label": label, title: label, type: "button" },
    cls: "clickable-icon slide-deck-viewer__button",
  });
  setIcon(button, icon);
  return button;
}

function renderMessage(parent: HTMLElement, message: string): void {
  parent.createDiv({ cls: "slide-deck-viewer__message", text: message });
}

function isPptxFile(file: TFile | null): file is TFile {
  return Boolean(file && PPTX_EXTENSIONS.includes(file.extension.toLowerCase()));
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
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

