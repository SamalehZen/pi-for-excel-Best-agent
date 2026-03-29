import type { DataAnalysisHints } from "./template-catalog.js";

const ANALYSIS_TAGS = [
  ["Columns", "Headers", "Data types", "Formulas", "Formatting"],
  ["Row patterns", "Date ranges", "Numeric fields", "Totals", "Categories"],
  ["Merged cells", "Borders", "Fonts", "Colors", "Alignment"],
];

function createTag(label: string, active: boolean): HTMLElement {
  const tag = document.createElement("span");
  tag.className = active ? "analysis-tag analysis-tag--active" : "analysis-tag";

  const icon = document.createElement("span");
  icon.className = "analysis-tag__icon";
  tag.appendChild(icon);

  const text = document.createTextNode(label);
  tag.appendChild(text);
  return tag;
}

function createTrackRow(tags: string[], reverse: boolean, slow: boolean, activeSet: Set<string>): HTMLElement {
  const track = document.createElement("div");
  let cls = "analysis-track";
  if (reverse) cls += " analysis-track--reverse";
  if (slow) cls += " analysis-track--slow";
  track.className = cls;

  const allTags = [...tags, ...tags, ...tags];
  for (const label of allTags) {
    track.appendChild(createTag(label, activeSet.has(label)));
  }

  return track;
}

export function renderAnalysisView(
  root: HTMLElement,
  hints: DataAnalysisHints | null,
  onComplete: () => void,
): void {
  root.innerHTML = "";

  const activeSet = new Set<string>();
  if (hints) {
    if (hints.hasDateColumns) activeSet.add("Date ranges");
    if (hints.hasNumericColumns) { activeSet.add("Numeric fields"); activeSet.add("Totals"); }
    if (hints.columnCount > 0) { activeSet.add("Columns"); activeSet.add("Headers"); }
    activeSet.add("Data types");
    activeSet.add("Row patterns");
  }

  const container = document.createElement("div");
  container.className = "analysis-container";

  const card = document.createElement("div");
  card.className = "analysis-card";

  const viewport = document.createElement("div");
  viewport.className = "analysis-viewport";

  for (let i = 0; i < ANALYSIS_TAGS.length; i++) {
    viewport.appendChild(createTrackRow(ANALYSIS_TAGS[i], i % 2 !== 0, i === 2, activeSet));
  }

  const lens = document.createElement("div");
  lens.className = "analysis-lens";
  viewport.appendChild(lens);

  const fadeLeft = document.createElement("div");
  fadeLeft.className = "analysis-fade-left";
  viewport.appendChild(fadeLeft);

  const fadeRight = document.createElement("div");
  fadeRight.className = "analysis-fade-right";
  viewport.appendChild(fadeRight);

  card.appendChild(viewport);

  const body = document.createElement("div");
  body.className = "analysis-body";

  const title = document.createElement("h3");
  title.className = "analysis-title";
  title.textContent = "Analyzing Your Data";
  body.appendChild(title);

  const desc = document.createElement("p");
  desc.className = "analysis-desc";
  desc.textContent = "Scanning structure, patterns, and formatting to find the best matching templates\u2026";
  body.appendChild(desc);

  const progress = document.createElement("div");
  progress.className = "analysis-progress";
  const bar = document.createElement("div");
  bar.className = "analysis-progress__bar";
  progress.appendChild(bar);
  body.appendChild(progress);

  if (hints) {
    const stats = document.createElement("div");
    stats.className = "analysis-stats";

    const statItems: [string, string][] = [
      [String(hints.columnCount), "Columns"],
      [String(hints.rowCount), "Rows"],
      [String(hints.headers.length), "Headers"],
    ];

    for (const [val, label] of statItems) {
      const stat = document.createElement("div");
      stat.className = "analysis-stat";
      const v = document.createElement("span");
      v.className = "analysis-stat__value";
      v.textContent = val;
      const l = document.createElement("span");
      l.className = "analysis-stat__label";
      l.textContent = label;
      stat.appendChild(v);
      stat.appendChild(l);
      stats.appendChild(stat);
    }
    body.appendChild(stats);
  }

  card.appendChild(body);
  container.appendChild(card);
  root.appendChild(container);

  setTimeout(() => {
    onComplete();
  }, 2800);
}
