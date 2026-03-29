export type {
  TemplatePalette,
  TemplateTypography,
  TemplateZoneType,
  TemplateZone,
  TemplateMetaField,
  TemplateColumn,
  TemplateSampleRow,
  TemplateStructure,
  TemplateSourceKind,
  TemplateDefinition,
  TemplateSummary,
} from "./types.js";

export {
  listTemplates,
  getTemplateById,
  getTemplatesByCategory,
  toTemplateSummary,
} from "./registry.js";

export {
  loadUserTemplates,
  saveUserTemplates,
  addUserTemplate,
  removeUserTemplate,
  USER_TEMPLATES_STORAGE_KEY,
  type TemplateSettingsStore,
} from "./store.js";

export { BUNDLED_TEMPLATE_DEFINITIONS } from "./definitions/index.js";
