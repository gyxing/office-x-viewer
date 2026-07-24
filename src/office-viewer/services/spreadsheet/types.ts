import type { OfficeChartModel } from '../../shared/ooxml/charts';

export type SpreadsheetWarning = {
  code: string;
  message: string;
  sheetName?: string;
  offset?: number;
};

export type SpreadsheetWorkbook = {
  sheets: SpreadsheetSheet[];
  warnings?: SpreadsheetWarning[];
  resources?: SpreadsheetResources;
};

export type SpreadsheetResources = {
  objectUrls: string[];
};

/** 释放工作簿创建的 Blob URL；重复调用保持幂等。 */
export function disposeSpreadsheetWorkbook(
  workbook: SpreadsheetWorkbook | undefined,
) {
  const urls = workbook?.resources?.objectUrls;
  if (!urls?.length) return;
  const uniqueUrls = new Set(urls);
  urls.length = 0;
  if (typeof URL === 'undefined' || typeof URL.revokeObjectURL !== 'function') {
    return;
  }
  uniqueUrls.forEach((url) => URL.revokeObjectURL(url));
}

export type SpreadsheetSheet = {
  id: string;
  name: string;
  path: string;
  kind?: 'worksheet' | 'chart';
  range?: string;
  rowCount: number;
  columnCount: number;
  columns: SpreadsheetColumn[];
  rows: SpreadsheetRow[];
  merges: SpreadsheetMerge[];
  images: SpreadsheetImage[];
  charts: SpreadsheetChart[];
};

export type SpreadsheetColumn = {
  index: number;
  label: string;
  width: number;
  hidden?: boolean;
};

export type SpreadsheetRow = {
  index: number;
  height: number;
  hidden?: boolean;
  cells: SpreadsheetCell[];
};

export type SpreadsheetCell = {
  ref: string;
  rowIndex: number;
  columnIndex: number;
  value: string;
  rawValue?: string;
  type?: string;
  styleId?: number;
  style?: SpreadsheetCellStyle;
  formula?: string;
  formulaTokens?: string;
  colSpan?: number;
  rowSpan?: number;
  hiddenByMerge?: boolean;
};

export type SpreadsheetMerge = {
  ref: string;
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
};

export type SpreadsheetAnchorPoint = {
  row: number;
  column: number;
  rowOffset: number;
  columnOffset: number;
};

export type SpreadsheetImage = {
  id: string;
  name?: string;
  src: string;
  alt?: string;
  from: SpreadsheetAnchorPoint;
  to: SpreadsheetAnchorPoint;
  x: number;
  y: number;
  width: number;
  height: number;
};

export type SpreadsheetChart = {
  id: string;
  title?: string;
  chart: OfficeChartModel;
  from: SpreadsheetAnchorPoint;
  to: SpreadsheetAnchorPoint;
  x: number;
  y: number;
  width: number;
  height: number;
};

export type SpreadsheetCellStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontFamily?: string;
  fontSize?: number;
  backgroundColor?: string;
  horizontalAlign?: 'left' | 'center' | 'right' | 'justify';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  border?: boolean;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderColor?: string;
  borderWidth?: number;
};
