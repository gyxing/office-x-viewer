// office-viewer 模块的公共入口，业务侧通过这里使用 OfficeViewer 及相关类型。
export { OfficeViewer } from './OfficeViewer';
export type { OfficeViewerProps, OfficeViewerUri } from './OfficeViewer';
export type { ParsedOfficeFile, PreviewKind } from './services/preview';
export { createOfficeParseSession } from './services/parsing';
export type {
  OfficeParseOptions,
  OfficeParseSession,
  OfficeParseSessionStatus,
  ParseProgress,
  ParseStage,
  WorkerMode,
} from './services/parsing';
export { parsePpt } from './services/ppt';
export { disposeDocDocument } from './services/doc/types';
export type {
  DocDocument,
  DocResources,
} from './services/doc/types';
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
