export const GALLERY_CHANNEL = "pi.template-gallery.v1";

export const GALLERY_ALLOWED_ORIGINS = new Set([
  "https://localhost:3000",
  "http://localhost:3000",
  "https://hyperexcel.vercel.app",
]);

export type GalleryMessageDirection = "host_to_gallery" | "gallery_to_host";

export interface GalleryAnalyzeRequest {
  channel: typeof GALLERY_CHANNEL;
  direction: "host_to_gallery";
  kind: "analyze";
  hints: {
    keywords: string[];
    hasDateColumns: boolean;
    hasNumericColumns: boolean;
    columnCount: number;
    rowCount: number;
    headers: string[];
  };
}

export interface GalleryShowRequest {
  channel: typeof GALLERY_CHANNEL;
  direction: "host_to_gallery";
  kind: "show";
  recommendedIds: string[];
}

export interface GalleryDismissRequest {
  channel: typeof GALLERY_CHANNEL;
  direction: "host_to_gallery";
  kind: "dismiss";
}

export type HostToGalleryMessage =
  | GalleryAnalyzeRequest
  | GalleryShowRequest
  | GalleryDismissRequest;

export interface GalleryReadyEvent {
  channel: typeof GALLERY_CHANNEL;
  direction: "gallery_to_host";
  kind: "ready";
}

export interface GalleryTemplateSelectedEvent {
  channel: typeof GALLERY_CHANNEL;
  direction: "gallery_to_host";
  kind: "template_selected";
  templateId: string;
  xlsxFile: string;
  templateName: string;
}

export interface GalleryClosedEvent {
  channel: typeof GALLERY_CHANNEL;
  direction: "gallery_to_host";
  kind: "closed";
}

export interface GalleryAnalysisDoneEvent {
  channel: typeof GALLERY_CHANNEL;
  direction: "gallery_to_host";
  kind: "analysis_done";
  recommendedIds: string[];
}

export type GalleryToHostMessage =
  | GalleryReadyEvent
  | GalleryTemplateSelectedEvent
  | GalleryClosedEvent
  | GalleryAnalysisDoneEvent;

export function isHostToGalleryMessage(data: unknown): data is HostToGalleryMessage {
  if (typeof data !== "object" || data === null) return false;
  const msg = data as Record<string, unknown>;
  return msg.channel === GALLERY_CHANNEL && msg.direction === "host_to_gallery";
}

export function isGalleryToHostMessage(data: unknown): data is GalleryToHostMessage {
  if (typeof data !== "object" || data === null) return false;
  const msg = data as Record<string, unknown>;
  return msg.channel === GALLERY_CHANNEL && msg.direction === "gallery_to_host";
}

function resolveOwnOrigin(): string {
  try {
    return window.location.origin;
  } catch {
    return "";
  }
}

export function isAllowedGalleryOrigin(origin: string): boolean {
  if (GALLERY_ALLOWED_ORIGINS.has(origin)) return true;
  const own = resolveOwnOrigin();
  return own.length > 0 && origin === own;
}
