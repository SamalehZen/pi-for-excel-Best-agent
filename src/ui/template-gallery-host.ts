import {
  GALLERY_CHANNEL,
  isGalleryToHostMessage,
  type HostToGalleryMessage,
  type GalleryToHostMessage,
} from "../template-gallery/bridge.js";
import { TEMPLATE_GALLERY_OVERLAY_ID } from "./overlay-ids.js";
import type { DataAnalysisHints } from "../template-gallery/template-catalog.js";

interface GalleryHostOptions {
  onTemplateSelected: (templateId: string, xlsxFile: string, templateName: string) => void;
  onClosed?: () => void;
  hints?: DataAnalysisHints | null;
  recommendedIds?: string[];
}

let activeOverlay: HTMLDivElement | null = null;
let activeIframe: HTMLIFrameElement | null = null;
let pendingMessage: HostToGalleryMessage | null = null;

function sendToGallery(message: HostToGalleryMessage): void {
  if (activeIframe?.contentWindow) {
    activeIframe.contentWindow.postMessage(message, "*");
  } else {
    pendingMessage = message;
  }
}

function removeGalleryOverlay(): void {
  if (activeOverlay) {
    activeOverlay.remove();
    activeOverlay = null;
  }
  activeIframe = null;
  pendingMessage = null;
}

export function isTemplateGalleryOpen(): boolean {
  return activeOverlay !== null && activeOverlay.isConnected;
}

export function dismissTemplateGallery(): void {
  if (activeIframe?.contentWindow) {
    sendToGallery({
      channel: GALLERY_CHANNEL,
      direction: "host_to_gallery",
      kind: "dismiss",
    });
  }
  removeGalleryOverlay();
}

export function showTemplateGallery(options: GalleryHostOptions): void {
  if (isTemplateGalleryOpen()) {
    dismissTemplateGallery();
  }

  const overlay = document.createElement("div");
  overlay.id = TEMPLATE_GALLERY_OVERLAY_ID;
  overlay.style.cssText = `
    position: fixed;
    inset: 0;
    z-index: 250;
    background: rgba(0, 0, 0, 0.4);
    display: flex;
    align-items: stretch;
    justify-content: stretch;
  `;

  const iframe = document.createElement("iframe");
  iframe.src = "/src/template-gallery.html";
  iframe.style.cssText = `
    width: 100%;
    height: 100%;
    border: none;
    background: #f7f7f5;
  `;
  iframe.setAttribute("title", "Template Gallery");

  overlay.appendChild(iframe);
  activeOverlay = overlay;
  activeIframe = iframe;

  const messageHandler = (event: MessageEvent): void => {
    if (!isGalleryToHostMessage(event.data)) return;
    const msg: GalleryToHostMessage = event.data;

    switch (msg.kind) {
      case "ready": {
        if (pendingMessage) {
          sendToGallery(pendingMessage);
          pendingMessage = null;
        } else if (options.hints) {
          sendToGallery({
            channel: GALLERY_CHANNEL,
            direction: "host_to_gallery",
            kind: "analyze",
            hints: options.hints,
          });
        } else if (options.recommendedIds) {
          sendToGallery({
            channel: GALLERY_CHANNEL,
            direction: "host_to_gallery",
            kind: "show",
            recommendedIds: options.recommendedIds,
          });
        } else {
          sendToGallery({
            channel: GALLERY_CHANNEL,
            direction: "host_to_gallery",
            kind: "show",
            recommendedIds: [],
          });
        }
        break;
      }

      case "template_selected":
        options.onTemplateSelected(msg.templateId, msg.xlsxFile, msg.templateName);
        removeGalleryOverlay();
        window.removeEventListener("message", messageHandler);
        break;

      case "closed":
        removeGalleryOverlay();
        window.removeEventListener("message", messageHandler);
        options.onClosed?.();
        break;

      case "analysis_done":
        break;
    }
  };

  window.addEventListener("message", messageHandler);

  const onEscape = (e: KeyboardEvent): void => {
    if (e.key === "Escape" && isTemplateGalleryOpen()) {
      e.preventDefault();
      dismissTemplateGallery();
      window.removeEventListener("message", messageHandler);
      window.removeEventListener("keydown", onEscape);
      options.onClosed?.();
    }
  };
  window.addEventListener("keydown", onEscape);

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) {
      dismissTemplateGallery();
      window.removeEventListener("message", messageHandler);
      window.removeEventListener("keydown", onEscape);
      options.onClosed?.();
    }
  });

  document.body.appendChild(overlay);
}
