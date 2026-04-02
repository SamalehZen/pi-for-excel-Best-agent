/**
 * Workbook blueprint — structural overview injected into model context.
 *
 * Re-uses buildOverview() from the get_workbook_overview tool but
 * wraps it with workbook-aware caching and invalidation signals.
 */

import { getWorkbookContext } from "../workbook/context.js";
import { buildOverview } from "../tools/get-workbook-overview.js";

const UNKNOWN_WORKBOOK_KEY = "__workbook_unknown__";

const blueprintCacheByWorkbook = new Map<string, string>();
const blueprintRevisionByWorkbook = new Map<string, number>();

function normalizeWorkbookKey(workbookId: string | null): string {
  return workbookId ?? UNKNOWN_WORKBOOK_KEY;
}

function getRevisionForKey(workbookKey: string): number {
  return blueprintRevisionByWorkbook.get(workbookKey) ?? 0;
}

function bumpRevisionForKey(workbookKey: string): number {
  const next = getRevisionForKey(workbookKey) + 1;
  blueprintRevisionByWorkbook.set(workbookKey, next);
  return next;
}

async function resolveWorkbookId(): Promise<string | null> {
  try {
    const ctx = await getWorkbookContext();
    return ctx.workbookId;
  } catch {
    return null;
  }
}

/**
 * Monotonic revision token for one workbook's blueprint cache.
 *
 * Intended for context injection logic: if this value changes for the active
 * workbook, workbook structure context should be considered stale.
 */
export function getBlueprintRevision(workbookId: string | null): number {
  return getRevisionForKey(normalizeWorkbookKey(workbookId));
}

/** Return the cached blueprint synchronously when already available. */
export function peekBlueprint(workbookId: string | null): string | undefined {
  return blueprintCacheByWorkbook.get(normalizeWorkbookKey(workbookId));
}

/** Get the workbook blueprint (cached per workbook identity when available). */
export async function getBlueprint(workbookId?: string | null): Promise<string> {
  const resolvedWorkbookId = workbookId === undefined
    ? await resolveWorkbookId()
    : workbookId;
  const workbookKey = normalizeWorkbookKey(resolvedWorkbookId);

  const cached = blueprintCacheByWorkbook.get(workbookKey);
  if (cached !== undefined) {
    return cached;
  }

  const blueprint = await buildOverview();
  blueprintCacheByWorkbook.set(workbookKey, blueprint);
  bumpRevisionForKey(workbookKey);
  return blueprint;
}

/** Force a fresh blueprint rebuild (e.g. after structural changes). */
export async function refreshBlueprint(workbookId?: string | null): Promise<string> {
  const resolvedWorkbookId = workbookId === undefined
    ? await resolveWorkbookId()
    : workbookId;
  invalidateBlueprint(resolvedWorkbookId);
  return getBlueprint(resolvedWorkbookId);
}

/** Invalidate one workbook's cached blueprint. */
export function invalidateBlueprint(workbookId: string | null): void {
  const workbookKey = normalizeWorkbookKey(workbookId);
  blueprintCacheByWorkbook.delete(workbookKey);
  bumpRevisionForKey(workbookKey);
}
