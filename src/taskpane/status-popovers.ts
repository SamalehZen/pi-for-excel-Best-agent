/**
 * Status-bar popovers (thinking selector + context quick actions).
 */

import type { ThinkingLevel } from "@mariozechner/pi-agent-core";

import type { StatusContextWarningSeverity } from "./status-context.js";

export type StatusCommandName = "compact" | "new";

type StatusPopoverKind = "thinking" | "context" | "template";

interface ActivePopoverState {
  kind: StatusPopoverKind;
  anchor: Element;
  popover: HTMLDivElement;
  cleanup: () => void;
}

interface ThinkingPopoverOptions {
  anchor: Element;
  description: string;
  levels: readonly ThinkingLevel[];
  activeLevel: ThinkingLevel;
  onSelectLevel: (level: ThinkingLevel) => void;
}

interface ContextPopoverOptions {
  anchor: Element;
  description: string;
  tokenDetail?: string;
  warning?: { text: string; severity: StatusContextWarningSeverity };
  onRunCommand: (command: StatusCommandName) => void;
}

const THINKING_LEVEL_LABELS: Record<ThinkingLevel, string> = {
  off: "Off",
  minimal: "Minimal",
  low: "Low",
  medium: "Medium",
  high: "High",
  xhigh: "Max",
};

const THINKING_LEVEL_HINTS: Record<ThinkingLevel, string> = {
  off: "Fastest — no reasoning step",
  minimal: "Quick — light reasoning",
  low: "Fast — moderate reasoning",
  medium: "Balanced — solid reasoning",
  high: "Slow — thorough reasoning",
  xhigh: "Slowest — deepest reasoning",
};

let activePopover: ActivePopoverState | null = null;

function clamp(value: number, min: number, max: number): number {
  if (max < min) return min;
  return Math.min(Math.max(value, min), max);
}

function normalizeDescription(text: string): string {
  const compact = text.replace(/\s+/g, " ").trim();
  return compact.length > 0 ? compact : "No details available.";
}

function createPopoverBase(kind: StatusPopoverKind): HTMLDivElement {
  const popover = document.createElement("div");
  popover.className = `pi-status-popover pi-status-popover--${kind}`;
  popover.setAttribute("role", "dialog");
  popover.setAttribute("aria-live", "polite");
  return popover;
}

function positionPopover(popover: HTMLDivElement, anchor: Element): void {
  const anchorRect = anchor.getBoundingClientRect();
  const viewportWidth = document.documentElement.clientWidth;
  const viewportHeight = document.documentElement.clientHeight;

  popover.style.left = "0px";
  popover.style.top = "0px";
  popover.style.visibility = "hidden";

  const popoverWidth = popover.offsetWidth;
  const popoverHeight = popover.offsetHeight;

  const maxLeft = Math.max(8, viewportWidth - popoverWidth - 8);
  const left = clamp(anchorRect.right - popoverWidth, 8, maxLeft);

  const preferredTop = anchorRect.top - popoverHeight - 8;
  const fallbackTop = anchorRect.bottom + 8;
  const maxTop = Math.max(8, viewportHeight - popoverHeight - 8);
  const top = preferredTop >= 8
    ? preferredTop
    : clamp(fallbackTop, 8, maxTop);

  popover.style.left = `${left}px`;
  popover.style.top = `${top}px`;
  popover.style.visibility = "visible";
}

function closeIfOpen(state: ActivePopoverState | null): void {
  if (!state) return;
  state.cleanup();
  state.popover.remove();
}

export function closeStatusPopover(): void {
  const state = activePopover;
  activePopover = null;
  closeIfOpen(state);
}

function mountPopover(kind: StatusPopoverKind, anchor: Element, popover: HTMLDivElement): void {
  closeStatusPopover();

  document.body.appendChild(popover);
  positionPopover(popover, anchor);

  const reposition = () => {
    if (!activePopover || activePopover.popover !== popover) return;
    positionPopover(popover, anchor);
  };

  const onMouseDown = (event: MouseEvent) => {
    const target = event.target;
    if (!(target instanceof Node)) return;

    if (popover.contains(target)) return;
    if (anchor instanceof Node && anchor.contains(target)) return;

    closeStatusPopover();
  };

  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key !== "Escape") return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    closeStatusPopover();
  };

  document.addEventListener("mousedown", onMouseDown, true);
  window.addEventListener("keydown", onKeyDown, true);
  window.addEventListener("resize", reposition);
  window.addEventListener("scroll", reposition, true);

  activePopover = {
    kind,
    anchor,
    popover,
    cleanup: () => {
      document.removeEventListener("mousedown", onMouseDown, true);
      window.removeEventListener("keydown", onKeyDown, true);
      window.removeEventListener("resize", reposition);
      window.removeEventListener("scroll", reposition, true);
    },
  };
}

function shouldToggle(kind: StatusPopoverKind, anchor: Element): boolean {
  return Boolean(
    activePopover
    && activePopover.kind === kind
    && activePopover.anchor === anchor,
  );
}

function createDescriptionBlock(text: string): HTMLParagraphElement {
  const description = document.createElement("p");
  description.className = "pi-status-popover__description";
  description.textContent = normalizeDescription(text);
  return description;
}

export function toggleThinkingPopover(opts: ThinkingPopoverOptions): void {
  if (shouldToggle("thinking", opts.anchor)) {
    closeStatusPopover();
    return;
  }

  const popover = createPopoverBase("thinking");

  const title = document.createElement("h3");
  title.className = "pi-status-popover__title";
  title.textContent = "Thinking level";

  const description = createDescriptionBlock(opts.description);

  const list = document.createElement("div");
  list.className = "pi-status-popover__list";

  for (const level of opts.levels) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "pi-status-popover__item";
    if (level === opts.activeLevel) {
      button.classList.add("is-active");
    }

    const body = document.createElement("span");
    body.className = "pi-status-popover__item-body";

    const label = document.createElement("span");
    label.className = "pi-status-popover__item-label";
    label.textContent = THINKING_LEVEL_LABELS[level];

    const hint = document.createElement("span");
    hint.className = "pi-status-popover__item-hint";
    hint.textContent = THINKING_LEVEL_HINTS[level];

    body.append(label, hint);

    const marker = document.createElement("span");
    marker.className = "pi-status-popover__item-marker";
    marker.textContent = level === opts.activeLevel ? "✓" : "";

    button.append(body, marker);

    button.addEventListener("click", () => {
      opts.onSelectLevel(level);
      closeStatusPopover();
    });

    list.appendChild(button);
  }

  popover.append(title, description, list);
  mountPopover("thinking", opts.anchor, popover);
}

function createCommandButton(args: {
  command: StatusCommandName;
  title: string;
  description: string;
  onRun: (command: StatusCommandName) => void;
}): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "pi-status-popover__command";

  const command = document.createElement("span");
  command.className = "pi-status-popover__command-name";
  command.textContent = `/${args.command}`;

  const description = document.createElement("span");
  description.className = "pi-status-popover__command-desc";
  description.textContent = args.description;

  button.append(command, description);
  button.setAttribute("aria-label", `${args.title}: ${args.description}`);

  button.addEventListener("click", () => {
    args.onRun(args.command);
    closeStatusPopover();
  });

  return button;
}

export interface TemplatePopoverItem {
  id: string;
  name: string;
  category: string;
  /** Preview thumbnail URL */
  previewUrl: string;
  /** Primary brand color for the color strip */
  primaryColor: string;
  /** Whether this template is recommended for the current sheet */
  isRecommended: boolean;
}

interface TemplatePopoverOptions {
  anchor: Element;
  templates: readonly TemplatePopoverItem[];
  /** Show loading indicator while analyzing sheet */
  loading?: boolean;
  onSelectTemplate: (templateId: string, mode: "full" | "design_only") => void;
}

export function toggleContextPopover(opts: ContextPopoverOptions): void {
  if (shouldToggle("context", opts.anchor)) {
    closeStatusPopover();
    return;
  }

  const popover = createPopoverBase("context");

  const title = document.createElement("h3");
  title.className = "pi-status-popover__title";
  title.textContent = "Context usage";

  const description = createDescriptionBlock(opts.description);

  if (opts.tokenDetail) {
    const detail = document.createElement("p");
    detail.className = "pi-status-popover__token-detail";
    detail.textContent = opts.tokenDetail;
    popover.append(title, description, detail);
  } else {
    popover.append(title, description);
  }

  if (opts.warning) {
    const warn = document.createElement("p");
    warn.className = `pi-status-popover__warning pi-status-popover__warning--${opts.warning.severity}`;
    warn.textContent = opts.warning.text;
    popover.appendChild(warn);
  }

  const actions = document.createElement("div");
  actions.className = "pi-status-popover__commands";

  actions.append(
    createCommandButton({
      command: "compact",
      title: "Compact conversation",
      description: "Summarize earlier messages to free space.",
      onRun: opts.onRunCommand,
    }),
    createCommandButton({
      command: "new",
      title: "Start new chat",
      description: "Open a fresh tab with empty context.",
      onRun: opts.onRunCommand,
    }),
  );

  popover.appendChild(actions);
  mountPopover("context", opts.anchor, popover);
}

/** Reference to the active template popover's list container for async updates. */
let activeTemplateListEl: HTMLDivElement | null = null;
let activeTemplateLoadingEl: HTMLParagraphElement | null = null;
let activeTemplateListView: HTMLDivElement | null = null;
let activeTemplateModeView: HTMLDivElement | null = null;
let activeTemplateOnSelect: ((id: string, mode: "full" | "design_only") => void) | null = null;

/**
 * Called from init.ts after async sheet analysis completes.
 * Updates the popover list with new recommended templates, or just removes loading state.
 */
export function updateTemplatePopoverList(newTemplates: TemplatePopoverItem[] | null): void {
  if (activeTemplateLoadingEl) {
    activeTemplateLoadingEl.remove();
    activeTemplateLoadingEl = null;
  }

  if (!newTemplates || !activeTemplateListEl || !activeTemplateOnSelect) return;

  activeTemplateListEl.innerHTML = "";
  renderTemplateCards(activeTemplateListEl, newTemplates, activeTemplateOnSelect);
}

function renderTemplateCards(
  list: HTMLDivElement,
  templates: readonly TemplatePopoverItem[],
  onSelect: (id: string, mode: "full" | "design_only") => void,
): void {
  for (const template of templates) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "pi-status-popover__item pi-template-card";

    const img = document.createElement("img");
    img.className = "pi-template-card__img";
    img.src = template.previewUrl;
    img.alt = template.name;
    img.loading = "lazy";

    const colorStrip = document.createElement("span");
    colorStrip.className = "pi-template-card__color-strip";
    colorStrip.style.background = template.primaryColor;

    const textWrap = document.createElement("span");
    textWrap.className = "pi-template-card__text";

    const label = document.createElement("span");
    label.className = "pi-status-popover__item-label";
    label.textContent = template.name;

    if (template.isRecommended) {
      const badge = document.createElement("span");
      badge.className = "pi-template-card__badge";
      badge.textContent = "\u2B50 Recommended";
      label.appendChild(document.createTextNode(" "));
      label.appendChild(badge);
    }

    const hint = document.createElement("span");
    hint.className = "pi-status-popover__item-hint";
    hint.textContent = template.category;

    textWrap.append(label, hint);
    button.append(img, colorStrip, textWrap);

    button.addEventListener("click", () => {
      showModeChoice(template.id, template.name, onSelect);
    });

    list.appendChild(button);
  }
}

function showModeChoice(
  templateId: string,
  templateName: string,
  onSelect: (id: string, mode: "full" | "design_only") => void,
): void {
  if (!activeTemplateListView || !activeTemplateModeView) return;

  activeTemplateListView.style.display = "none";
  activeTemplateModeView.style.display = "";
  activeTemplateModeView.innerHTML = "";

  const backBtn = document.createElement("button");
  backBtn.type = "button";
  backBtn.className = "pi-template-back";
  backBtn.textContent = "\u2190 Back";
  backBtn.addEventListener("click", () => {
    if (activeTemplateModeView) activeTemplateModeView.style.display = "none";
    if (activeTemplateListView) activeTemplateListView.style.display = "";
  });

  const modeTitle = document.createElement("h3");
  modeTitle.className = "pi-status-popover__title";
  modeTitle.textContent = templateName;

  const modeDesc = createDescriptionBlock("How do you want to apply this template?");

  const modeList = document.createElement("div");
  modeList.className = "pi-status-popover__list";

  const fullBtn = document.createElement("button");
  fullBtn.type = "button";
  fullBtn.className = "pi-status-popover__item pi-template-mode-btn";
  const fullBody = document.createElement("span");
  fullBody.className = "pi-status-popover__item-body";
  const fullLabel = document.createElement("span");
  fullLabel.className = "pi-status-popover__item-label";
  fullLabel.textContent = "\uD83D\uDCCB Full Template";
  const fullHint = document.createElement("span");
  fullHint.className = "pi-status-popover__item-hint";
  fullHint.textContent = "Replace sheet with complete template structure + sample data";
  fullBody.append(fullLabel, fullHint);
  fullBtn.appendChild(fullBody);
  fullBtn.addEventListener("click", () => {
    onSelect(templateId, "full");
    closeStatusPopover();
  });

  const designBtn = document.createElement("button");
  designBtn.type = "button";
  designBtn.className = "pi-status-popover__item pi-template-mode-btn";
  const designBody = document.createElement("span");
  designBody.className = "pi-status-popover__item-body";
  const designLabel = document.createElement("span");
  designLabel.className = "pi-status-popover__item-label";
  designLabel.textContent = "\uD83C\uDFA8 Design Only";
  const designHint = document.createElement("span");
  designHint.className = "pi-status-popover__item-hint";
  designHint.textContent = "Apply colors, fonts & formatting to your existing data";
  designBody.append(designLabel, designHint);
  designBtn.appendChild(designBody);
  designBtn.addEventListener("click", () => {
    onSelect(templateId, "design_only");
    closeStatusPopover();
  });

  modeList.append(fullBtn, designBtn);
  activeTemplateModeView.append(backBtn, modeTitle, modeDesc, modeList);
}

export function toggleTemplatePopover(opts: TemplatePopoverOptions): void {
  if (shouldToggle("template", opts.anchor)) {
    closeStatusPopover();
    return;
  }

  const popover = createPopoverBase("template");

  activeTemplateOnSelect = opts.onSelectTemplate;

  const listView = document.createElement("div");
  listView.className = "pi-template-view";
  activeTemplateListView = listView;

  const title = document.createElement("h3");
  title.className = "pi-status-popover__title";
  title.textContent = "Template";

  const description = createDescriptionBlock("Choose a design template to apply to the active sheet.");

  let loadingEl: HTMLParagraphElement | null = null;
  if (opts.loading) {
    loadingEl = document.createElement("p");
    loadingEl.className = "pi-template-loading";
    loadingEl.textContent = "\u2728 Analyzing your sheet\u2026";
    activeTemplateLoadingEl = loadingEl;
  }

  const list = document.createElement("div");
  list.className = "pi-status-popover__list";
  activeTemplateListEl = list;

  renderTemplateCards(list, opts.templates, opts.onSelectTemplate);

  listView.append(title, description);
  if (loadingEl) listView.appendChild(loadingEl);
  listView.appendChild(list);

  const modeView = document.createElement("div");
  modeView.className = "pi-template-view";
  modeView.style.display = "none";
  activeTemplateModeView = modeView;

  popover.append(listView, modeView);

  const origCleanup = (): void => {
    activeTemplateListEl = null;
    activeTemplateLoadingEl = null;
    activeTemplateListView = null;
    activeTemplateModeView = null;
    activeTemplateOnSelect = null;
  };

  mountPopover("template", opts.anchor, popover);

  if (activePopover) {
    const existingCleanup = activePopover.cleanup;
    activePopover.cleanup = () => {
      existingCleanup();
      origCleanup();
    };
  }
}

