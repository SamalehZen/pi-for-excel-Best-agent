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
