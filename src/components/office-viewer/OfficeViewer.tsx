import { Layout } from 'antd';
import { memo, useMemo } from 'react';
import type { DocxDocument } from '../../services/docx/types';
import type { PptxDocument } from '../../services/pptx/types';
import type { PreviewKind } from '../../services/officePreview';
import type { XlsxWorkbook } from '../../services/xlsx/types';
import { DocxViewer } from '../docx-viewer/DocxViewer';
import { PptxViewer } from '../pptx-viewer/PptxViewer';
import { XlsxViewer } from '../xlsx-viewer/XlsxViewer';
import { OfficeError } from './OfficeError';
import { OfficeLoading } from './OfficeLoading';
import { OfficeToolbar } from './OfficeToolbar';

const { Content } = Layout;

type OfficeViewerProps = {
  fileName: string;
  loading: boolean;
  error?: string;
  previewKind: PreviewKind;
  pptxDocument?: PptxDocument;
  xlsxWorkbook?: XlsxWorkbook;
  docxDocument?: DocxDocument;
  activeIndex: number;
  activeSheetId?: string;
  zoom: number;
  onSelectSlide: (index: number) => void;
  onSelectSheet: (sheetId: string) => void;
  onZoomChange: (zoom: number) => void;
  onUpload: (file: File) => void;
};

function OfficeViewerComponent({
  fileName,
  loading,
  error,
  previewKind,
  pptxDocument,
  xlsxWorkbook,
  docxDocument,
  activeIndex,
  activeSheetId,
  zoom,
  onSelectSlide,
  onSelectSheet,
  onZoomChange,
  onUpload,
}: OfficeViewerProps) {
  const hasDocument = useMemo(
    () =>
      previewKind === 'pptx'
        ? Boolean(pptxDocument?.slides.length)
        : previewKind === 'xlsx'
          ? Boolean(xlsxWorkbook?.sheets.length)
          : Boolean(docxDocument?.blocks.length),
    [docxDocument, pptxDocument, previewKind, xlsxWorkbook],
  );

  const canGoPreviousSlide = previewKind === 'pptx' && Boolean(pptxDocument?.slides.length) && activeIndex > 0;
  const canGoNextSlide =
    previewKind === 'pptx' &&
    Boolean(pptxDocument?.slides.length) &&
    activeIndex < (pptxDocument?.slides.length ?? 1) - 1;

  const handlePreviousSlide = () => {
    onSelectSlide(Math.max(activeIndex - 1, 0));
  };

  const handleNextSlide = () => {
    onSelectSlide(Math.min(activeIndex + 1, (pptxDocument?.slides.length ?? 1) - 1));
  };

  const handleZoomOut = () => {
    onZoomChange(Math.max(25, zoom - 25));
  };

  const handleZoomIn = () => {
    onZoomChange(Math.min(300, zoom + 25));
  };

  const handleResetZoom = () => {
    onZoomChange(100);
  };

  return (
    <Layout style={{ minHeight: '100vh', background: '#eef1f6' }}>
      <OfficeToolbar
        fileName={fileName}
        previewKind={previewKind}
        zoom={zoom}
        hasDocument={hasDocument}
        canGoPreviousSlide={canGoPreviousSlide}
        canGoNextSlide={canGoNextSlide}
        onUpload={onUpload}
        onPreviousSlide={handlePreviousSlide}
        onNextSlide={handleNextSlide}
        onZoomOut={handleZoomOut}
        onZoomIn={handleZoomIn}
        onZoomChange={onZoomChange}
        onResetZoom={handleResetZoom}
      />
      <Content style={{ background: '#eef1f6', height: 'calc(100vh - 56px)', overflow: 'hidden' }}>
        {error ? (
          <OfficeError message={error} />
        ) : loading ? (
          <OfficeLoading />
        ) : previewKind === 'xlsx' ? (
          <XlsxViewer
            workbook={xlsxWorkbook}
            activeSheetId={activeSheetId}
            zoom={zoom}
            onSelectSheet={onSelectSheet}
          />
        ) : previewKind === 'docx' ? (
          <DocxViewer document={docxDocument} zoom={zoom} />
        ) : (
          <PptxViewer
            document={pptxDocument}
            activeIndex={activeIndex}
            zoom={zoom}
            onSelectSlide={onSelectSlide}
          />
        )}
      </Content>
    </Layout>
  );
}

export const OfficeViewer = memo(OfficeViewerComponent);
