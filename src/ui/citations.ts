import { qualifiedAddress } from "../excel/helpers.js";
import { navigateToAddress, navigateToWorksheet } from "./cell-link.js";

const CITE_PREFIX = "#cite:";

export interface CiteTarget {
  sheet: string;
  range?: string;
}

function decodeCitePart(part: string): string {
  try {
    return decodeURIComponent(part);
  } catch {
    return part;
  }
}

export function isCiteHref(href: string): boolean {
  return href.startsWith(CITE_PREFIX);
}

export function parseCiteHref(href: string): CiteTarget | null {
  if (!isCiteHref(href)) return null;

  const payload = href.slice(CITE_PREFIX.length);
  if (payload.length === 0) return null;

  const bangIndex = payload.indexOf("!");
  const rawSheet = bangIndex >= 0 ? payload.slice(0, bangIndex) : payload;
  const rawRange = bangIndex >= 0 ? payload.slice(bangIndex + 1) : undefined;
  const sheet = decodeCitePart(rawSheet).trim();
  const range = rawRange === undefined ? undefined : decodeCitePart(rawRange).trim();

  if (sheet.length === 0) return null;
  if (rawRange !== undefined && !range) return null;

  return range ? { sheet, range } : { sheet };
}

export function findCiteAnchor(event: Event): HTMLAnchorElement | null {
  for (const node of event.composedPath()) {
    if (!(node instanceof HTMLAnchorElement)) continue;

    const href = node.getAttribute("href");
    if (href && isCiteHref(href)) {
      return node;
    }
  }

  return null;
}

export async function navigateToCitation(target: CiteTarget): Promise<void> {
  if (target.range) {
    await navigateToAddress(qualifiedAddress(target.sheet, target.range));
    return;
  }

  await navigateToWorksheet(target.sheet);
}
