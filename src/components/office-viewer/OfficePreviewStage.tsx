// OfficePreviewStage 根据当前文件格式切换到对应预览组件，并统一处理加载和错误态。
import { memo } from 'react';
import type { DocDocument } from '../../services/doc/types';
import type { DocxDocument } from '../../services/docx/types';
import type { PreviewKind } from '../../services/office/preview';
import type { PptxDocument } from '../../services/pptx/types';
import type { XlsxWorkbook } from '../../services/xlsx/types';
import { DocViewer } from '../doc-viewer/DocViewer';
import { DocxViewer } from '../docx-viewer/DocxViewer';
import { PptxViewer } from '../pptx-viewer/PptxViewer';
import { XlsxViewer } from '../xlsx-viewer/XlsxViewer';
import { OfficeError } from './OfficeError';
import { OfficeLoading } from './OfficeLoading';

type OfficePreviewStageProps = {
  loading: boolean;
  error?: string;
  previewKind: PreviewKind;
  pptxDocument?: PptxDocument;
  xlsxWorkbook?: XlsxWorkbook;
  docxDocument?: DocxDocument;
  docDocument?: DocDocument;
  activeIndex: number;
  activeSheetId?: string;
  zoom: number;
  onSelectSlide: (index: number) => void;
  onSelectSheet: (sheetId: string) => void;
};

function OfficePreviewStageComponent({
  loading,
  error,
  previewKind,
  pptxDocument,
  xlsxWorkbook,
  docxDocument,
  docDocument,
  activeIndex,
  activeSheetId,
  zoom,
  onSelectSlide,
  onSelectSheet,
}: OfficePreviewStageProps) {
  if (error) return <OfficeError message={error} />;
  if (loading) return <OfficeLoading />;

  // 渲染分发只保留格式选择逻辑，各格式 viewer 自己处理空文档和格式内部状态。
  if (previewKind === 'xlsx') {
    return <XlsxViewer workbook={xlsxWorkbook} activeSheetId={activeSheetId} zoom={zoom} onSelectSheet={onSelectSheet} />;
  }

  if (previewKind === 'docx') {
    return <DocxViewer document={docxDocument} zoom={zoom} />;
  }

  if (previewKind === 'doc') {
    return <DocViewer document={docDocument} zoom={zoom} />;
  }

  return <PptxViewer document={pptxDocument} activeIndex={activeIndex} zoom={zoom} onSelectSlide={onSelectSlide} />;
}

export const OfficePreviewStage = memo(OfficePreviewStageComponent);
