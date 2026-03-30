import "./gallery.css";
import { renderAnalysisView } from "./analysis-view.js";
import { renderGalleryView } from "./gallery-view.js";
import { findRecommendedTemplates, type DataAnalysisHints } from "./template-catalog.js";
import {
  GALLERY_CHANNEL,
  isHostToGalleryMessage,
  isAllowedGalleryOrigin,
  type GalleryToHostMessage,
} from "./bridge.js";

const galleryRoot = document.getElementById("gallery-root");
if (!galleryRoot) throw new Error("Missing #gallery-root");
const root: HTMLElement = galleryRoot;

// Show loading state immediately
root.innerHTML = `
  <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;font-family:var(--gallery-font,'DM Sans',sans-serif);color:#71717a;gap:12px;">
    <div style="width:24px;height:24px;border:2px solid rgba(0,0,0,0.1);border-top-color:#71717a;border-radius:50%;animation:spin 0.8s linear infinite;"></div>
    <p style="font-size:13px;margin:0;">Loading templates\u2026</p>
  </div>
  <style>@keyframes spin { to { transform: rotate(360deg); } }</style>
`;

// Fallback: if no host message received within 5s, show error
const fallbackTimeout = setTimeout(() => {
  if (root.querySelector('.gallery-container') === null && root.querySelector('.analysis-container') === null) {
    root.innerHTML = `
      <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;font-family:var(--gallery-font,'DM Sans',sans-serif);color:#71717a;gap:12px;padding:24px;text-align:center;">
        <p style="font-size:15px;font-weight:600;color:#1a1a1a;margin:0;">Unable to load gallery</p>
        <p style="font-size:13px;margin:0;">The template gallery could not connect to the add-in.</p>
      </div>
    `;
  }
}, 5000);

let currentRecommendedIds: string[] = [];

const hostOrigin: string = (() => {
  try {
    if (window.parent && window.parent !== window) {
      return window.location.origin;
    }
  } catch {
    // cross-origin parent — fall back to own origin
  }
  return window.location.origin;
})();

function sendToHost(message: GalleryToHostMessage): void {
  if (window.parent && window.parent !== window) {
    window.parent.postMessage(message, hostOrigin);
  }
}

function showAnalysis(hints: DataAnalysisHints | null): void {
  renderAnalysisView(root, hints, () => {
    const recommendations = hints ? findRecommendedTemplates(hints) : [];
    currentRecommendedIds = recommendations.map((r) => r.id);

    sendToHost({
      channel: GALLERY_CHANNEL,
      direction: "gallery_to_host",
      kind: "analysis_done",
      recommendedIds: currentRecommendedIds,
    });

    showGallery(currentRecommendedIds);
  });
}

function showGallery(recommendedIds: string[]): void {
  root.innerHTML = "";
  renderGalleryView(
    root,
    recommendedIds,
    (template) => {
      sendToHost({
        channel: GALLERY_CHANNEL,
        direction: "gallery_to_host",
        kind: "template_selected",
        templateId: template.id,
        xlsxFile: template.xlsxFile,
        templateName: template.name,
      });
    },
    () => {
      sendToHost({
        channel: GALLERY_CHANNEL,
        direction: "gallery_to_host",
        kind: "closed",
      });
    },
  );
}

window.addEventListener("message", (event: MessageEvent) => {
  clearTimeout(fallbackTimeout);
  if (typeof event.origin === "string" && !isAllowedGalleryOrigin(event.origin)) return;
  if (!isHostToGalleryMessage(event.data)) return;

  const msg = event.data;

  switch (msg.kind) {
    case "analyze":
      showAnalysis(msg.hints);
      break;

    case "show":
      currentRecommendedIds = msg.recommendedIds;
      showGallery(msg.recommendedIds);
      break;

    case "dismiss":
      root.innerHTML = "";
      break;
  }
});

sendToHost({
  channel: GALLERY_CHANNEL,
  direction: "gallery_to_host",
  kind: "ready",
});

const params = new URLSearchParams(window.location.search);
if (params.get("standalone") === "1") {
  showAnalysis({
    keywords: ["sales", "report"],
    hasDateColumns: true,
    hasNumericColumns: true,
    columnCount: 8,
    rowCount: 50,
    headers: ["Date", "Product", "Amount", "Quantity", "Total"],
  });
}
