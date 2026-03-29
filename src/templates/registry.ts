/**
 * Template registry — manages bundled and user-added design templates.
 */

import type { TemplateDefinition, TemplateSummary } from "./types.js";
import { BUNDLED_TEMPLATE_DEFINITIONS } from "./definitions/index.js";

/** Get a summary for a TemplateDefinition. */
export function toTemplateSummary(def: TemplateDefinition): TemplateSummary {
  return {
    id: def.id,
    name: def.name,
    category: def.category,
    description: def.description,
    sourceKind: def.sourceKind,
    primaryColor: def.design.palette.headerBg,
    fontFamily: def.design.typography.fontFamily,
    columnCount: def.structure.columns.length,
  };
}

/** List all available templates (bundled + user). */
export function listTemplates(userTemplates?: readonly TemplateDefinition[]): TemplateSummary[] {
  const all: TemplateDefinition[] = [...BUNDLED_TEMPLATE_DEFINITIONS];
  if (userTemplates) {
    for (const ut of userTemplates) {
      const idx = all.findIndex((t) => t.id === ut.id);
      if (idx >= 0) all[idx] = ut;
      else all.push(ut);
    }
  }
  return all.map(toTemplateSummary);
}

/** Get a template by ID. Returns null if not found. */
export function getTemplateById(
  id: string,
  userTemplates?: readonly TemplateDefinition[],
): TemplateDefinition | null {
  const needle = id.trim().toLowerCase();
  if (userTemplates) {
    const userMatch = userTemplates.find((t) => t.id.toLowerCase() === needle);
    if (userMatch) return userMatch;
  }
  const bundled = BUNDLED_TEMPLATE_DEFINITIONS.find(
    (t) => t.id.toLowerCase() === needle,
  );
  return bundled ?? null;
}

/** Get templates by category. */
export function getTemplatesByCategory(
  category: string,
  userTemplates?: readonly TemplateDefinition[],
): TemplateSummary[] {
  return listTemplates(userTemplates).filter(
    (t) => t.category.toLowerCase() === category.toLowerCase(),
  );
}
