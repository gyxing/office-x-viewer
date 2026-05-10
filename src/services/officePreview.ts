import type { DocxDocument } from './docx/types';
import type { PptxDocument } from './pptx/types';
import type { XlsxWorkbook } from './xlsx/types';
import { parseDocx } from './docx/parseDocx';
import { parsePptx } from './pptx/parsePptx';
import { parseXlsx } from './xlsx/parseXlsx';

export type PreviewKind = 'pptx' | 'xlsx' | 'docx';

export type ParsedOfficeFile =
  | { kind: 'pptx'; document: PptxDocument }
  | { kind: 'xlsx'; workbook: XlsxWorkbook }
  | { kind: 'docx'; document: DocxDocument };

export function detectPreviewKind(fileName: string): PreviewKind {
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.xlsx')) return 'xlsx';
  if (lower.endsWith('.docx')) return 'docx';
  return 'pptx';
}

export async function parseOfficeFile(file: File): Promise<ParsedOfficeFile> {
  const kind = detectPreviewKind(file.name);

  if (kind === 'xlsx') {
    return { kind, workbook: await parseXlsx(file) };
  }

  if (kind === 'docx') {
    return { kind, document: await parseDocx(file) };
  }

  return { kind, document: await parsePptx(file) };
}
