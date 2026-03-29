/**
 * User template persistence via SettingsStore.
 *
 * User templates are stored per-workbook to allow workbook-specific design libraries.
 */

import type { TemplateDefinition } from "./types.js";
import { isRecord } from "../utils/type-guards.js";

export const USER_TEMPLATES_STORAGE_KEY = "templates.user.v1";

interface StoredUserTemplatesDocument {
  version: 1;
  templates: TemplateDefinition[];
}

export interface TemplateSettingsStore {
  get: (key: string) => Promise<unknown>;
  set: (key: string, value: unknown) => Promise<void>;
}

function isValidStoredTemplate(t: unknown): t is TemplateDefinition {
  if (!isRecord(t)) return false;
  if (typeof t.id !== "string" || typeof t.name !== "string") return false;
  if (typeof t.category !== "string" || typeof t.description !== "string") return false;
  if (t.sourceKind !== "bundled" && t.sourceKind !== "user") return false;

  if (!isRecord(t.design)) return false;
  if (typeof t.design.alternatingRows !== "boolean" || typeof t.design.titleBold !== "boolean") return false;

  if (!isRecord(t.design.palette)) return false;
  const p = t.design.palette;
  if (
    typeof p.titleBg !== "string" || typeof p.titleFg !== "string" ||
    typeof p.headerBg !== "string" || typeof p.headerFg !== "string" ||
    typeof p.labelBg !== "string" || typeof p.labelFg !== "string" ||
    typeof p.accentBg !== "string" || typeof p.accentFg !== "string" ||
    typeof p.alternateBg !== "string" ||
    typeof p.totalBg !== "string" || typeof p.totalFg !== "string"
  ) return false;

  if (!isRecord(t.design.typography)) return false;
  const ty = t.design.typography;
  if (
    typeof ty.fontFamily !== "string" ||
    typeof ty.titleSize !== "number" || typeof ty.sectionHeaderSize !== "number" ||
    typeof ty.headerSize !== "number" || typeof ty.bodySize !== "number"
  ) return false;

  if (!isRecord(t.structure)) return false;
  const s = t.structure;
  if (typeof s.title !== "string" || typeof s.titleRow !== "number") return false;
  if (typeof s.headerRow !== "number" || typeof s.columnSpan !== "string") return false;
  if (typeof s.totalRows !== "number") return false;
  if (!Array.isArray(s.columns) || !Array.isArray(s.metaFields)) return false;
  if (!Array.isArray(s.sampleData) || !Array.isArray(s.zones)) return false;

  return true;
}

function parseStoredTemplates(raw: unknown): TemplateDefinition[] {
  if (!isRecord(raw) || raw.version !== 1) return [];
  if (!Array.isArray(raw.templates)) return [];
  return (raw.templates as unknown[]).filter(isValidStoredTemplate);
}

export async function loadUserTemplates(
  settings: TemplateSettingsStore,
): Promise<TemplateDefinition[]> {
  const raw = await settings.get(USER_TEMPLATES_STORAGE_KEY);
  return parseStoredTemplates(raw);
}

export async function saveUserTemplates(
  settings: TemplateSettingsStore,
  templates: TemplateDefinition[],
): Promise<void> {
  const doc: StoredUserTemplatesDocument = { version: 1, templates };
  await settings.set(USER_TEMPLATES_STORAGE_KEY, doc);
}

export async function addUserTemplate(
  settings: TemplateSettingsStore,
  template: TemplateDefinition,
): Promise<void> {
  const existing = await loadUserTemplates(settings);
  const needle = template.id.toLowerCase();
  const idx = existing.findIndex((t) => t.id.toLowerCase() === needle);
  if (idx >= 0) existing[idx] = template;
  else existing.push(template);
  await saveUserTemplates(settings, existing);
}

export async function removeUserTemplate(
  settings: TemplateSettingsStore,
  templateId: string,
): Promise<boolean> {
  const existing = await loadUserTemplates(settings);
  const needle = templateId.toLowerCase();
  const idx = existing.findIndex((t) => t.id.toLowerCase() === needle);
  if (idx < 0) return false;
  existing.splice(idx, 1);
  await saveUserTemplates(settings, existing);
  return true;
}
