export const OFFICEJS_API_DOCS_PATH = "/home/user/docs/excel-officejs-api.d.ts";

export function getOfficeJsDocsContent(): string {
  return `// Excel Office.js API Reference (curated subset)
// Full docs: https://learn.microsoft.com/en-us/javascript/api/excel

declare namespace Excel {
  interface ClientResult<T> {
    value: T;
  }

  interface Workbook {
    worksheets: WorksheetCollection;
    names: NamedItemCollection;
    tables: TableCollection;
  }

  interface WorksheetCollection {
    getItem(key: string): Worksheet;
    getActiveWorksheet(): Worksheet;
    items: Worksheet[];
  }

  interface Worksheet {
    name: string;
    id: string;
    charts: ChartCollection;
    tables: TableCollection;
    pivotTables: PivotTableCollection;
    getRange(address?: string): Range;
    getUsedRange(valuesOnly?: boolean): Range;
    getUsedRangeOrNullObject(valuesOnly?: boolean): Range;
    activate(): void;
    delete(): void;
    load(properties: string): void;
  }

  interface Range {
    address: string;
    values: unknown[][];
    formulas: unknown[][];
    text: string[][];
    numberFormat: string[][];
    rowIndex: number;
    columnIndex: number;
    rowCount: number;
    columnCount: number;
    format: RangeFormat;
    dataValidation: DataValidation;
    getImage(): ClientResult<string>;
    getColumn(column: number): Range;
    getRow(row: number): Range;
    merge(across?: boolean): void;
    unmerge(): void;
    copyFrom(source: Range, copyType?: RangeCopyType): void;
    delete(shift: DeleteShiftDirection): void;
    load(properties: string): void;
  }

  interface RangeFormat {
    autofitColumns(): void;
    autofitRows(): void;
  }

  interface ChartCollection {
    add(type: ChartType, sourceData: Range, seriesBy?: ChartSeriesBy): Chart;
    getItem(name: string): Chart;
    getCount(): ClientResult<number>;
    items: Chart[];
  }

  interface Chart {
    id: string;
    name: string;
    chartType: ChartType;
    title: ChartTitle;
    axes: ChartAxes;
    legend: ChartLegend;
    dataLabels: ChartDataLabels;
    width: number;
    height: number;
    setPosition(startCell: Range, endCell?: Range): void;
    setData(sourceData: Range, seriesBy?: ChartSeriesBy): void;
    delete(): void;
  }

  interface ChartTitle {
    text: string;
  }

  interface ChartAxes {
    categoryAxis: ChartAxis;
    valueAxis: ChartAxis;
  }

  interface ChartAxis {
    title: ChartAxisTitle;
  }

  interface ChartAxisTitle {
    text: string;
  }

  interface ChartLegend {
    visible: boolean;
  }

  interface ChartDataLabels {
    showValue: boolean;
  }

  enum ChartType {
    columnClustered,
    barClustered,
    line,
    pie,
    xyscatter,
    area,
    doughnut,
    radar,
  }

  enum ChartSeriesBy {
    auto,
    columns,
    rows,
  }

  interface PivotTableCollection {
    add(name: string, source: Range, destination: Range): PivotTable;
    getItem(name: string): PivotTable;
    items: PivotTable[];
  }

  interface PivotTable {
    id: string;
    name: string;
    hierarchies: PivotHierarchyCollection;
    rowHierarchies: RowColumnPivotHierarchyCollection;
    columnHierarchies: RowColumnPivotHierarchyCollection;
    dataHierarchies: DataPivotHierarchyCollection;
    filterHierarchies: FilterPivotHierarchyCollection;
    delete(): void;
  }

  interface PivotHierarchyCollection {
    items: PivotHierarchy[];
  }

  interface PivotHierarchy {
    name: string;
  }

  interface RowColumnPivotHierarchyCollection {
    add(hierarchy: PivotHierarchy): void;
  }

  interface DataPivotHierarchyCollection {
    add(hierarchy: PivotHierarchy): DataPivotHierarchy;
  }

  interface DataPivotHierarchy {
    summarizeBy: AggregationFunction;
  }

  interface FilterPivotHierarchyCollection {
    add(hierarchy: PivotHierarchy): void;
  }

  enum AggregationFunction {
    sum,
    count,
    average,
    max,
    min,
  }

  interface TableCollection {
    add(address: Range | string, hasHeaders: boolean): Table;
    getItem(name: string): Table;
    items: Table[];
  }

  interface Table {
    name: string;
    style: string;
    showHeaders: boolean;
    showTotals: boolean;
    columns: TableColumnCollection;
    rows: TableRowCollection;
    delete(): void;
  }

  interface TableColumnCollection {
    items: TableColumn[];
  }

  interface TableColumn {
    name: string;
  }

  interface TableRowCollection {
    items: TableRow[];
  }

  interface TableRow {
    index: number;
  }

  interface NamedItemCollection {
    getItem(name: string): NamedItem;
    items: NamedItem[];
  }

  interface NamedItem {
    name: string;
    type: string;
  }

  interface DataValidation {
    rule: DataValidationRule;
    prompt: DataValidationPrompt;
    errorAlert: DataValidationErrorAlert;
    clear(): void;
  }

  interface DataValidationRule {
    custom?: {
      formula: string;
    };
    list?: {
      inCellDropDown: boolean;
      source: string;
    };
  }

  interface DataValidationPrompt {
    showPrompt: boolean;
    title: string;
    message: string;
  }

  interface DataValidationErrorAlert {
    showAlert: boolean;
    title: string;
    message: string;
  }

  enum DeleteShiftDirection {
    up,
    left,
  }

  enum RangeCopyType {
    all,
    formulas,
    values,
    formats,
  }

  enum DataValidationOperator {
    between,
    notBetween,
    equalTo,
    notEqualTo,
    greaterThan,
    lessThan,
    greaterThanOrEqualTo,
    lessThanOrEqualTo,
  }
}
`;
}
