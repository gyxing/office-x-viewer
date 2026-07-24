import type { DocDocument } from './doc/types';
import type { DocxDocument } from './docx/types';
import type { PptxDocument } from './pptx/types';
import type { PresentationDocument } from './presentation/types';
import { createOfficeParseSession } from './parsing';
import type { OfficeParseOptions } from './parsing';
import type { SpreadsheetWorkbook } from './spreadsheet/types';

// 组件入口只关心“文件类型 + 解析结果”，具体格式的包结构解析都收敛在各自 service 中。
export type { PreviewKind } from './parsing/detectPreviewKind';
export {
  detectPreviewKind,
  isPresentationPreviewKind,
  isSpreadsheetPreviewKind,
  isSupportedOfficeFileName,
  SUPPORTED_OFFICE_EXTENSIONS,
} from './parsing/detectPreviewKind';

export type ParsedOfficeFile =
  | { kind: 'pptx'; document: PptxDocument }
  | { kind: 'ppt'; document: PresentationDocument }
  | { kind: 'xlsx'; workbook: SpreadsheetWorkbook }
  | { kind: 'xls'; workbook: SpreadsheetWorkbook }
  | { kind: 'docx'; document: DocxDocument }
  | { kind: 'doc'; document: DocDocument };

export async function parseOfficeFile(
  file: File,
  options?: OfficeParseOptions,
): Promise<ParsedOfficeFile> {
  const session = createOfficeParseSession(file, options);
  try {
    return await session.result;
  } finally {
    session.dispose();
  }
}
