import type { DocDocument } from './doc/types';
import type { DocxDocument } from './docx/types';
import type { PptxDocument } from './pptx/types';
import type { XlsxWorkbook } from './xlsx/types';
import { parseDoc } from './doc/parseDoc';
import { parseDocx } from './docx/parseDocx';
import { parsePptx } from './pptx/parsePptx';
import { parseXlsx } from './xlsx/parseXlsx';

export type PreviewKind = 'pptx' | 'xlsx' | 'docx' | 'doc';

export type ParsedOfficeFile =
  | { kind: 'pptx'; document: PptxDocument }
  | { kind: 'xlsx'; workbook: XlsxWorkbook }
  | { kind: 'docx'; document: DocxDocument }
  | { kind: 'doc'; document: DocDocument };

export function detectPreviewKind(fileName: string): PreviewKind {
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.xlsx')) return 'xlsx';
  if (lower.endsWith('.docx')) return 'docx';
  if (lower.endsWith('.doc')) return 'doc';
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

  if (kind === 'doc') {
    return { kind, document: await parseDoc(file) };
  }

  return { kind, document: await parsePptx(file) };
}
