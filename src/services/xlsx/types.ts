export type XlsxWorkbook = {
  sheets: XlsxSheet[];
};

export type XlsxSheet = {
  id: string;
  name: string;
  path: string;
  range?: string;
  rowCount: number;
  columnCount: number;
  columns: XlsxColumn[];
  rows: XlsxRow[];
  merges: XlsxMerge[];
  images: XlsxImage[];
  charts: XlsxChart[];
};

export type XlsxColumn = {
  index: number;
  label: string;
  width: number;
  hidden?: boolean;
};

export type XlsxRow = {
  index: number;
  height: number;
  cells: XlsxCell[];
};

export type XlsxCell = {
  ref: string;
  rowIndex: number;
  columnIndex: number;
  value: string;
  rawValue?: string;
  type?: string;
  styleId?: number;
  style?: XlsxCellStyle;
  colSpan?: number;
  rowSpan?: number;
  hiddenByMerge?: boolean;
};

export type XlsxMerge = {
  ref: string;
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
};

export type XlsxImage = {
  id: string;
  name?: string;
  src: string;
  alt?: string;
  from: XlsxAnchorPoint;
  to: XlsxAnchorPoint;
  x: number;
  y: number;
  width: number;
  height: number;
};

export type XlsxChart = {
  id: string;
  title?: string;
  chart: import('../office/charts').OfficeChartModel;
  from: XlsxAnchorPoint;
  to: XlsxAnchorPoint;
  x: number;
  y: number;
  width: number;
  height: number;
};

export type XlsxAnchorPoint = {
  row: number;
  column: number;
  rowOffset: number;
  columnOffset: number;
};

export type XlsxCellStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  backgroundColor?: string;
  horizontalAlign?: 'left' | 'center' | 'right' | 'justify';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  border?: boolean;
};
