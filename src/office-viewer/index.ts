// office-viewer 模块的公共入口，业务侧通过这里使用 OfficeViewer 及相关类型。
export { OfficeViewer } from './OfficeViewer';
export type { OfficeViewerProps, OfficeViewerUri } from './OfficeViewer';
export type { ParsedOfficeFile, PreviewKind } from './services/preview';
export { parsePpt } from './services/ppt';
export { disposePresentationDocument } from './services/presentation/dispose';
export type { PresentationDocument } from './services/presentation/types';
export { disposeSpreadsheetWorkbook } from './services/spreadsheet/types';
export type {
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetResources,
  SpreadsheetSheet,
  SpreadsheetWarning,
  SpreadsheetWorkbook,
} from './services/spreadsheet/types';
