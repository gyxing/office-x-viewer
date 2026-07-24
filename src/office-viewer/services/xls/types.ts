import type { SpreadsheetWarning } from '../spreadsheet/types';
import type { Biff8Record } from './biff8/Biff8Reader';

export type Biff8SheetType =
  | 'worksheet'
  | 'chart'
  | 'macro'
  | 'dialog'
  | 'unknown';

export type Biff8SheetDescriptor = {
  id: string;
  name: string;
  streamOffset: number;
  visibility: 'visible' | 'hidden' | 'veryHidden';
  type: Biff8SheetType;
};

export type Biff8Font = {
  name: string;
  heightTwips: number;
  colorIndex: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
};

export type Biff8BorderStyle = {
  style: number;
  colorIndex: number;
};

export type Biff8CellFormat = {
  fontIndex: number;
  formatIndex: number;
  parentStyleIndex: number;
  isStyle: boolean;
  horizontalAlign?: number;
  verticalAlign?: number;
  wrapText?: boolean;
  fillPattern?: number;
  fillForegroundColorIndex?: number;
  fillBackgroundColorIndex?: number;
  leftBorder?: Biff8BorderStyle;
  rightBorder?: Biff8BorderStyle;
  topBorder?: Biff8BorderStyle;
  bottomBorder?: Biff8BorderStyle;
};

export type Biff8DefinedName = {
  id: number;
  name: string;
  tokens: Uint8Array;
};

export type Biff8WorkbookGlobals = {
  sheets: Biff8SheetDescriptor[];
  sharedStrings: string[];
  fonts: Biff8Font[];
  formats: Map<number, string>;
  cellFormats: Biff8CellFormat[];
  palette: string[];
  date1904: boolean;
  definedNames: Biff8DefinedName[];
  warnings: SpreadsheetWarning[];
  hasVba: boolean;
  codePage?: number;
  drawingGroupRecords: Biff8RecordSequence[];
};

export type Biff8RecordSequence = {
  recordId: number;
  offset: number;
  chunks: Uint8Array[];
};

export type Biff8ChartSubstream = {
  offset: number;
  records: Biff8Record[];
};

export type Biff8Cell = {
  row: number;
  column: number;
  xfIndex: number;
  value: string | number | boolean | null;
  cachedType: 'string' | 'number' | 'boolean' | 'error' | 'blank';
  formula?: string;
  formulaTokens?: string;
};

export type Biff8RowInfo = {
  index: number;
  heightTwips?: number;
  hidden?: boolean;
  outlineLevel?: number;
};

export type Biff8ColumnInfo = {
  firstColumn: number;
  lastColumn: number;
  widthCharacters: number;
  hidden?: boolean;
  outlineLevel?: number;
  xfIndex?: number;
};

export type Biff8Merge = {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
};

export type Biff8Worksheet = {
  descriptor: Biff8SheetDescriptor;
  cells: Biff8Cell[];
  rows: Biff8RowInfo[];
  columns: Biff8ColumnInfo[];
  merges: Biff8Merge[];
  defaultColumnWidth: number;
  defaultRowHeightTwips: number;
  dimensions?: {
    firstRow: number;
    lastRowExclusive: number;
    firstColumn: number;
    lastColumnExclusive: number;
  };
  hasDrawingRecords: boolean;
  hasChartRecords: boolean;
  chartSubstreams: Biff8ChartSubstream[];
  drawingRecords: Biff8RecordSequence[];
  warnings: SpreadsheetWarning[];
};

export type Biff8ChartSheet = {
  descriptor: Biff8SheetDescriptor;
  substream: Biff8ChartSubstream;
};

export type Biff8Workbook = {
  globals: Biff8WorkbookGlobals;
  worksheets: Biff8Worksheet[];
  chartSheets: Biff8ChartSheet[];
  warnings: SpreadsheetWarning[];
};
