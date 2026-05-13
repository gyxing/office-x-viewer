// OfficePreviewStage 根据当前文件格式切换到对应预览组件，并统一处理加载和错误态。
import { lazy, memo, Suspense } from 'react';
import type { DocDocument } from '../services/doc/types';
import type { DocxDocument } from '../services/docx/types';
import type { PreviewKind } from '../services/preview';
import type { PptxDocument } from '../services/pptx/types';
import type { XlsxWorkbook } from '../services/xlsx/types';
import { OfficeError } from './Error';
import { OfficeLoading } from './Loading';

const LazyPptxViewer = lazy(() => import('../formats/pptx/PptxViewer').then((module) => ({ default: module.PptxViewer })));
const LazyXlsxViewer = lazy(() => import('../formats/xlsx/XlsxViewer').then((module) => ({ default: module.XlsxViewer })));
const LazyDocxViewer = lazy(() => import('../formats/docx/DocxViewer').then((module) => ({ default: module.DocxViewer })));
const LazyDocViewer = lazy(() => import('../formats/doc/DocViewer').then((module) => ({ default: module.DocViewer })));

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

  // 格式 viewer 是真正的重渲染模块，按文件类型懒加载，避免首屏一次性拉取所有预览实现。
  return (
    <Suspense fallback={<OfficeLoading />}>
      {previewKind === 'xlsx' ? (
        <LazyXlsxViewer workbook={xlsxWorkbook} activeSheetId={activeSheetId} zoom={zoom} onSelectSheet={onSelectSheet} />
      ) : previewKind === 'docx' ? (
        <LazyDocxViewer document={docxDocument} zoom={zoom} />
      ) : previewKind === 'doc' ? (
        <LazyDocViewer document={docDocument} zoom={zoom} />
      ) : (
        <LazyPptxViewer document={pptxDocument} activeIndex={activeIndex} zoom={zoom} onSelectSlide={onSelectSlide} />
      )}
    </Suspense>
  );
}

export const OfficePreviewStage = memo(OfficePreviewStageComponent);
