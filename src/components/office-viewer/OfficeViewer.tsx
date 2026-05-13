// OfficeViewer 是组件库对外主入口，负责文件上传解析、格式状态和全局工具栏交互。
import { Layout } from 'antd';
import { memo, useCallback, useEffect, useMemo, useState } from 'react';
import type { CSSProperties } from 'react';
import type { DocDocument } from '../../services/doc/types';
import type { DocxDocument } from '../../services/docx/types';
import { detectPreviewKind, parseOfficeFile, type ParsedOfficeFile, type PreviewKind } from '../../services/office/preview';
import type { PptxDocument } from '../../services/pptx/types';
import type { XlsxWorkbook } from '../../services/xlsx/types';
import './index.less';
import { OfficePreviewStage } from './OfficePreviewStage';
import { OfficeToolbar } from './OfficeToolbar';
import { OFFICE_DEFAULT_ZOOM, OFFICE_MAX_ZOOM, OFFICE_MIN_ZOOM, OFFICE_ZOOM_STEP } from './shared/constants';

const { Content } = Layout;

type OfficeViewerProps = {
  initialFile?: File;
  defaultFileName?: string;
  defaultPreviewKind?: PreviewKind;
  defaultZoom?: number;
  uploadAccept?: string;
  uploadLabel?: string;
  className?: string;
  style?: CSSProperties;
  onFileParsed?: (parsed: ParsedOfficeFile, file: File) => void;
  onError?: (error: Error, file?: File) => void;
};

function OfficeViewerComponent({
  initialFile,
  defaultFileName = '未加载文件',
  defaultPreviewKind = 'pptx',
  defaultZoom = OFFICE_DEFAULT_ZOOM,
  uploadAccept,
  uploadLabel,
  className,
  style,
  onFileParsed,
  onError,
}: OfficeViewerProps) {
  // OfficeViewer 是公共组件入口，集中管理“文件状态”和“格式私有状态”，避免使用者再组合多个子组件。
  const [fileName, setFileName] = useState(defaultFileName);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>();
  const [previewKind, setPreviewKind] = useState<PreviewKind>(defaultPreviewKind);
  const [pptxDocument, setPptxDocument] = useState<PptxDocument>();
  const [xlsxWorkbook, setXlsxWorkbook] = useState<XlsxWorkbook>();
  const [docxDocument, setDocxDocument] = useState<DocxDocument>();
  const [docDocument, setDocDocument] = useState<DocDocument>();
  const [activeIndex, setActiveIndex] = useState(0);
  const [activeSheetId, setActiveSheetId] = useState<string>();
  const [zoom, setZoom] = useState(defaultZoom);

  const handleUpload = useCallback(
    async (file: File) => {
      setLoading(true);
      setError(undefined);

      try {
        // 上传新文件时同步重置所有格式相关状态，防止上一份文档的页码/缩放/工作表残留到新文档。
        const fileKind = detectPreviewKind(file.name);
        setPreviewKind(fileKind);
        setFileName(file.name);
        setActiveIndex(0);
        setZoom(defaultZoom);

        const parsed = await parseOfficeFile(file);
        setPptxDocument(parsed.kind === 'pptx' ? parsed.document : undefined);
        setXlsxWorkbook(parsed.kind === 'xlsx' ? parsed.workbook : undefined);
        setDocxDocument(parsed.kind === 'docx' ? parsed.document : undefined);
        setDocDocument(parsed.kind === 'doc' ? parsed.document : undefined);
        setActiveSheetId(parsed.kind === 'xlsx' ? parsed.workbook.sheets[0]?.id : undefined);
        onFileParsed?.(parsed, file);
      } catch (nextError) {
        // 对外回调始终给 Error 实例，组件内部只保存可展示的 message。
        const normalizedError = nextError instanceof Error ? nextError : new Error('文件解析失败');
        setError(normalizedError.message);
        onError?.(normalizedError, file);
      } finally {
        setLoading(false);
      }
    },
    [defaultZoom, onError, onFileParsed],
  );

  useEffect(() => {
    if (!initialFile) return;
    void handleUpload(initialFile);
  }, [handleUpload, initialFile]);

  const hasDocument = useMemo(
    () =>
      // 工具栏的翻页/缩放按钮只依赖“当前格式是否有可渲染内容”，不要耦合到具体 viewer 实现。
      previewKind === 'pptx'
        ? Boolean(pptxDocument?.slides.length)
        : previewKind === 'xlsx'
          ? Boolean(xlsxWorkbook?.sheets.length)
          : previewKind === 'docx'
            ? Boolean(docxDocument?.blocks.length)
            : Boolean(docDocument?.paragraphs.length),
    [docDocument, docxDocument, pptxDocument, previewKind, xlsxWorkbook],
  );

  const canGoPreviousSlide = previewKind === 'pptx' && Boolean(pptxDocument?.slides.length) && activeIndex > 0;
  const canGoNextSlide =
    previewKind === 'pptx' &&
    Boolean(pptxDocument?.slides.length) &&
    activeIndex < (pptxDocument?.slides.length ?? 1) - 1;

  const handlePreviousSlide = useCallback(() => {
    setActiveIndex((value) => Math.max(value - 1, 0));
  }, []);

  const handleNextSlide = useCallback(() => {
    setActiveIndex((value) => Math.min(value + 1, (pptxDocument?.slides.length ?? 1) - 1));
  }, [pptxDocument?.slides.length]);

  const handleZoomOut = useCallback(() => {
    setZoom((value) => Math.max(OFFICE_MIN_ZOOM, value - OFFICE_ZOOM_STEP));
  }, []);

  const handleZoomIn = useCallback(() => {
    setZoom((value) => Math.min(OFFICE_MAX_ZOOM, value + OFFICE_ZOOM_STEP));
  }, []);

  const handleResetZoom = useCallback(() => {
    setZoom(defaultZoom);
  }, [defaultZoom]);

  return (
    <Layout className={['oxv-office-viewer', className].filter(Boolean).join(' ')} style={style}>
      <OfficeToolbar
        fileName={fileName}
        previewKind={previewKind}
        uploadAccept={uploadAccept}
        uploadLabel={uploadLabel}
        zoom={zoom}
        hasDocument={hasDocument}
        canGoPreviousSlide={canGoPreviousSlide}
        canGoNextSlide={canGoNextSlide}
        onUpload={handleUpload}
        onPreviousSlide={handlePreviousSlide}
        onNextSlide={handleNextSlide}
        onZoomOut={handleZoomOut}
        onZoomIn={handleZoomIn}
        onZoomChange={setZoom}
        onResetZoom={handleResetZoom}
      />
      <Content className="oxv-office-viewer__content">
        <OfficePreviewStage
          loading={loading}
          error={error}
          previewKind={previewKind}
          pptxDocument={pptxDocument}
          xlsxWorkbook={xlsxWorkbook}
          docxDocument={docxDocument}
          docDocument={docDocument}
          activeIndex={activeIndex}
          activeSheetId={activeSheetId}
          zoom={zoom}
          onSelectSlide={setActiveIndex}
          onSelectSheet={setActiveSheetId}
        />
      </Content>
    </Layout>
  );
}

export const OfficeViewer = memo(OfficeViewerComponent);
