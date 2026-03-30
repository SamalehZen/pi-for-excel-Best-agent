import {
  GALLERY_CHANNEL,
  isGalleryToHostMessage,
  isAllowedGalleryOrigin,
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
let cleanupListeners: (() => void) | null = null;

function getIframeOrigin(): string {
  return window.location.origin;
}

function sendToGallery(message: HostToGalleryMessage): void {
  if (activeIframe?.contentWindow) {
    activeIframe.contentWindow.postMessage(message, getIframeOrigin());
  } else {
    pendingMessage = message;
  }
}

function teardown(): void {
  if (cleanupListeners) {
    cleanupListeners();
    cleanupListeners = null;
  }
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
  teardown();
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

  // Close button — always visible above the iframe
  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.setAttribute("aria-label", "Close gallery");
  closeBtn.textContent = "\u2715";
  closeBtn.style.cssText = `
    position: absolute;
    top: 8px;
    right: 8px;
    z-index: 260;
    width: 32px;
    height: 32px;
    border-radius: 50%;
    border: 1px solid rgba(255,255,255,0.3);
    background: rgba(0,0,0,0.5);
    color: #fff;
    font-size: 16px;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    backdrop-filter: blur(4px);
  `;
  closeBtn.addEventListener("click", () => closeGallery(true));
  overlay.appendChild(closeBtn);

  activeOverlay = overlay;
  activeIframe = iframe;

  const closeGallery = (notifyClosed: boolean): void => {
    teardown();
    if (notifyClosed) options.onClosed?.();
  };

  const messageHandler = (event: MessageEvent): void => {
    if (typeof event.origin === "string" && !isAllowedGalleryOrigin(event.origin)) return;
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
        teardown();
        break;

      case "closed":
        closeGallery(true);
        break;

      case "analysis_done":
        break;
    }
  };

  const onEscape = (e: KeyboardEvent): void => {
    if (e.key === "Escape" && isTemplateGalleryOpen()) {
      e.preventDefault();
      closeGallery(true);
    }
  };

  const onBackdropClick = (e: MouseEvent): void => {
    if (e.target === overlay) {
      closeGallery(true);
    }
  };

  window.addEventListener("message", messageHandler);
  window.addEventListener("keydown", onEscape);
  overlay.addEventListener("click", onBackdropClick);

  cleanupListeners = () => {
    window.removeEventListener("message", messageHandler);
    window.removeEventListener("keydown", onEscape);
    overlay.removeEventListener("click", onBackdropClick);
  };

  document.body.appendChild(overlay);
}
