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

function parseStoredTemplates(raw: unknown): TemplateDefinition[] {
  if (!isRecord(raw) || raw.version !== 1) return [];
  if (!Array.isArray(raw.templates)) return [];
  return (raw.templates as unknown[]).filter(
    (t): t is TemplateDefinition =>
      isRecord(t) && typeof t.id === "string" && typeof t.name === "string",
  );
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
  const idx = existing.findIndex((t) => t.id === template.id);
  if (idx >= 0) existing[idx] = template;
  else existing.push(template);
  await saveUserTemplates(settings, existing);
}

export async function removeUserTemplate(
  settings: TemplateSettingsStore,
  templateId: string,
): Promise<boolean> {
  const existing = await loadUserTemplates(settings);
  const idx = existing.findIndex((t) => t.id === templateId);
  if (idx < 0) return false;
  existing.splice(idx, 1);
  await saveUserTemplates(settings, existing);
  return true;
}
