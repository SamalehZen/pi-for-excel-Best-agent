import { TEMPLATE_CATALOG, type TemplateInfo } from "./template-catalog.js";

function getUniqueCategories(): string[] {
  const cats = new Set(TEMPLATE_CATALOG.map((t) => t.category));
  return Array.from(cats).sort();
}

function createPreviewCard(
  template: TemplateInfo,
  isRecommended: boolean,
  onClick: (t: TemplateInfo) => void,
): HTMLElement {
  const card = document.createElement("div");
  card.className = isRecommended
    ? "gallery-card gallery-card--recommended"
    : "gallery-card";
  card.setAttribute("role", "button");
  card.setAttribute("tabindex", "0");
  card.setAttribute("aria-label", `Select ${template.name} template`);

  if (isRecommended) {
    const badge = document.createElement("span");
    badge.className = "gallery-card__badge";
    badge.textContent = "Recommended";
    card.appendChild(badge);
  }

  const img = document.createElement("img");
  img.className = "gallery-card__preview";
  img.src = template.previewUrl;
  img.alt = `${template.name} preview`;
  img.loading = "lazy";
  img.onerror = () => {
    img.style.display = "none";
  };
  card.appendChild(img);

  const strip = document.createElement("div");
  strip.className = "gallery-card__color-strip";
  strip.style.background = template.primaryColor;
  card.appendChild(strip);

  const body = document.createElement("div");
  body.className = "gallery-card__body";

  const name = document.createElement("p");
  name.className = "gallery-card__name";
  name.textContent = template.name;
  body.appendChild(name);

  const category = document.createElement("p");
  category.className = "gallery-card__category";
  category.textContent = `${template.category} \u00B7 ${template.fontFamily}`;
  body.appendChild(category);

  card.appendChild(body);

  card.addEventListener("click", () => onClick(template));
  card.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      onClick(template);
    }
  });

  return card;
}

function createExpandedView(
  template: TemplateInfo,
  isRecommended: boolean,
  onApply: (t: TemplateInfo) => void,
  onClose: () => void,
): HTMLElement {
  const overlay = document.createElement("div");
  overlay.className = "gallery-expanded";

  const card = document.createElement("div");
  card.className = "gallery-expanded__card";

  const img = document.createElement("img");
  img.className = "gallery-expanded__preview";
  img.src = template.previewUrl;
  img.alt = `${template.name} preview`;
  img.onerror = () => {
    img.style.display = "none";
  };
  card.appendChild(img);

  const body = document.createElement("div");
  body.className = "gallery-expanded__body";

  const name = document.createElement("h2");
  name.className = "gallery-expanded__name";
  name.textContent = template.name;
  body.appendChild(name);

  const meta = document.createElement("div");
  meta.className = "gallery-expanded__meta";

  const catItem = document.createElement("span");
  catItem.className = "gallery-expanded__meta-item";
  catItem.textContent = template.category;
  meta.appendChild(catItem);

  const fontItem = document.createElement("span");
  fontItem.className = "gallery-expanded__meta-item";
  fontItem.textContent = template.fontFamily;
  meta.appendChild(fontItem);

  const colorItem = document.createElement("span");
  colorItem.className = "gallery-expanded__meta-item";
  const dot = document.createElement("span");
  dot.className = "gallery-expanded__meta-dot";
  dot.style.background = template.primaryColor;
  colorItem.appendChild(dot);
  colorItem.appendChild(document.createTextNode(template.primaryColor));
  meta.appendChild(colorItem);

  body.appendChild(meta);

  const desc = document.createElement("p");
  desc.className = "gallery-expanded__desc";
  desc.textContent = template.description;
  if (isRecommended) {
    desc.textContent += " \u2728 This template is recommended based on your data structure.";
  }
  body.appendChild(desc);

  const actions = document.createElement("div");
  actions.className = "gallery-expanded__actions";

  const cancelBtn = document.createElement("button");
  cancelBtn.className = "gallery-expanded__btn";
  cancelBtn.textContent = "Cancel";
  cancelBtn.type = "button";
  actions.appendChild(cancelBtn);

  const applyBtn = document.createElement("button");
  applyBtn.className = "gallery-expanded__btn gallery-expanded__btn--primary";
  applyBtn.textContent = "Apply Template";
  applyBtn.type = "button";
  actions.appendChild(applyBtn);

  body.appendChild(actions);
  card.appendChild(body);
  overlay.appendChild(card);

  const onEscape = (e: KeyboardEvent) => {
    if (e.key === "Escape") {
      e.preventDefault();
      document.removeEventListener("keydown", onEscape);
      onClose();
    }
  };
  document.addEventListener("keydown", onEscape);

  const cleanupAndClose = (): void => {
    document.removeEventListener("keydown", onEscape);
    onClose();
  };

  const cleanupAndApply = (): void => {
    document.removeEventListener("keydown", onEscape);
    onApply(template);
  };

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) cleanupAndClose();
  });

  cancelBtn.addEventListener("click", cleanupAndClose);
  applyBtn.addEventListener("click", cleanupAndApply);

  return overlay;
}

export function renderGalleryView(
  root: HTMLElement,
  recommendedIds: string[],
  onSelect: (template: TemplateInfo) => void,
  onClose: () => void,
): void {
  root.innerHTML = "";

  const container = document.createElement("div");
  container.className = "gallery-container view-enter";

  const header = document.createElement("div");
  header.className = "gallery-header";

  const backBtn = document.createElement("button");
  backBtn.className = "gallery-back";
  backBtn.type = "button";
  const arrow = document.createElement("span");
  arrow.className = "gallery-back__arrow";
  arrow.textContent = "\u2190";
  backBtn.appendChild(arrow);
  backBtn.appendChild(document.createTextNode("Close"));
  backBtn.addEventListener("click", onClose);
  header.appendChild(backBtn);

  const titleArea = document.createElement("div");
  titleArea.className = "gallery-title-area";
  const title = document.createElement("h2");
  title.className = "gallery-title";
  title.textContent = "Template Gallery";
  titleArea.appendChild(title);
  const subtitle = document.createElement("p");
  subtitle.className = "gallery-subtitle";
  subtitle.textContent = `${TEMPLATE_CATALOG.length} professional templates`;
  titleArea.appendChild(subtitle);
  header.appendChild(titleArea);

  container.appendChild(header);

  const categories = getUniqueCategories();
  const filterBar = document.createElement("div");
  filterBar.className = "gallery-filter-bar";

  const allFilter = document.createElement("button");
  allFilter.className = "gallery-filter gallery-filter--active";
  allFilter.textContent = "All";
  allFilter.type = "button";
  filterBar.appendChild(allFilter);

  for (const cat of categories) {
    const btn = document.createElement("button");
    btn.className = "gallery-filter";
    btn.textContent = cat;
    btn.type = "button";
    btn.dataset.category = cat;
    filterBar.appendChild(btn);
  }

  container.appendChild(filterBar);

  const grid = document.createElement("div");
  grid.className = "gallery-grid";

  const recommendedSet = new Set(recommendedIds);

  const sortedTemplates = [...TEMPLATE_CATALOG].sort((a, b) => {
    const aRec = recommendedSet.has(a.id) ? 0 : 1;
    const bRec = recommendedSet.has(b.id) ? 0 : 1;
    if (aRec !== bRec) return aRec - bRec;
    return a.name.localeCompare(b.name);
  });

  let expandedOverlay: HTMLElement | null = null;

  const renderCards = (filter: string | null): void => {
    grid.innerHTML = "";
    const filtered = filter
      ? sortedTemplates.filter((t) => t.category === filter)
      : sortedTemplates;

    for (const template of filtered) {
      const card = createPreviewCard(
        template,
        recommendedSet.has(template.id),
        (t) => {
          if (expandedOverlay) {
            expandedOverlay.remove();
            expandedOverlay = null;
          }
          expandedOverlay = createExpandedView(
            t,
            recommendedSet.has(t.id),
            (selected) => {
              if (expandedOverlay) {
                expandedOverlay.remove();
                expandedOverlay = null;
              }
              onSelect(selected);
            },
            () => {
              if (expandedOverlay) {
                expandedOverlay.remove();
                expandedOverlay = null;
              }
            },
          );
          document.body.appendChild(expandedOverlay);
        },
      );
      grid.appendChild(card);
    }
  };

  renderCards(null);
  container.appendChild(grid);

  filterBar.addEventListener("click", (e) => {
    const target = e.target as HTMLElement;
    if (!target.classList.contains("gallery-filter")) return;

    for (const btn of filterBar.querySelectorAll(".gallery-filter")) {
      btn.classList.remove("gallery-filter--active");
    }
    target.classList.add("gallery-filter--active");

    const cat = target.dataset.category ?? null;
    renderCards(cat);
  });

  root.appendChild(container);
}
