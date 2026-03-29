/**
 * Design template type definitions.
 *
 * A template captures a complete spreadsheet design: visual styling (palette,
 * typography), structural layout (zones, columns), and optional sample data.
 * Templates can be applied in two modes:
 * - "full": recreate structure + data + formatting on a blank sheet.
 * - "design_only": detect existing data layout and apply only visual formatting.
 */

/** Color palette extracted from a template. */
export interface TemplatePalette {
  /** Title bar background */
  titleBg: string;
  /** Title bar text color */
  titleFg: string;
  /** Column/section header background */
  headerBg: string;
  /** Column/section header text color */
  headerFg: string;
  /** Label/field name background */
  labelBg: string;
  /** Label/field name text color */
  labelFg: string;
  /** Accent/highlight background (e.g. computed cells, amounts) */
  accentBg: string;
  /** Accent text color */
  accentFg: string;
  /** Alternate data row background (empty string = no alternating) */
  alternateBg: string;
  /** Total/summary row background */
  totalBg: string;
  /** Total/summary row text color */
  totalFg: string;
}

/** Typography settings. */
export interface TemplateTypography {
  /** Primary body font family */
  fontFamily: string;
  /** Title font family (defaults to fontFamily if omitted) */
  titleFontFamily?: string;
  /** Title font size in points */
  titleSize: number;
  /** Section/category header font size */
  sectionHeaderSize: number;
  /** Column header font size */
  headerSize: number;
  /** Body text font size */
  bodySize: number;
}

/** Row zone types recognized in template layout. */
export type TemplateZoneType =
  | "title"
  | "meta"
  | "section_header"
  | "column_header"
  | "data"
  | "data_alternate"
  | "total"
  | "spacer"
  | "accent";

/** Defines which zone type applies to a row or row range. */
export interface TemplateZone {
  type: TemplateZoneType;
  /** 1-indexed row number or [startRow, endRow] inclusive range */
  rows: number | [number, number];
}

/** Metadata field in the template header area (e.g. "Employee:", "Date:"). */
export interface TemplateMetaField {
  label: string;
  placeholder: string;
  row: number;
  labelCol: string;
  valueCol: string;
}

/** Column definition for the data table area. */
export interface TemplateColumn {
  header: string;
  col: string;
  width?: number;
  alignment?: "Left" | "Center" | "Right" | "General";
  isAccent?: boolean;
  isBold?: boolean;
}

export type TemplateSampleRow = (string | number | null)[];

/** Complete template structure. */
export interface TemplateStructure {
  title: string;
  titleRow: number;
  metaFields: TemplateMetaField[];
  headerRow: number;
  columns: TemplateColumn[];
  sampleData: TemplateSampleRow[];
  totalRow?: {
    row: number;
    label?: string;
    labelCol?: string;
    values: Record<string, string | number>;
  };
  zones: TemplateZone[];
  columnSpan: string;
  totalRows: number;
}

export type TemplateSourceKind = "bundled" | "user";

/** Complete template definition. */
export interface TemplateDefinition {
  id: string;
  name: string;
  category: string;
  description: string;
  design: {
    palette: TemplatePalette;
    typography: TemplateTypography;
    alternatingRows: boolean;
    titleBold: boolean;
    titleRowHeight?: number;
    defaultRowHeight?: number;
    headerRowHeight?: number;
  };
  structure: TemplateStructure;
  sourceKind: TemplateSourceKind;
}

/** Lightweight template summary for list responses. */
export interface TemplateSummary {
  id: string;
  name: string;
  category: string;
  description: string;
  sourceKind: TemplateSourceKind;
  primaryColor: string;
  fontFamily: string;
  columnCount: number;
}
