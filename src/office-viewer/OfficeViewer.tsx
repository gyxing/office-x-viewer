// OfficeViewer 是组件库对外主入口，负责文件上传解析、格式状态和全局工具栏交互。
import { Layout } from 'antd';
import { memo, useCallback, useEffect, useMemo, useState } from 'react';
import type { CSSProperties } from 'react';
import type { DocDocument } from './services/doc/types';
import type { DocxDocument } from './services/docx/types';
import {
  detectPreviewKind,
  isSupportedOfficeFileName,
  parseOfficeFile,
  type ParsedOfficeFile,
  type PreviewKind,
} from './services/preview';
import type { PptxDocument } from './services/pptx/types';
import type { XlsxWorkbook } from './services/xlsx/types';
import './index.less';
import { OfficePreviewStage } from './shell/PreviewStage';
import { OfficeToolbar } from './shell/Toolbar';
import { OFFICE_DEFAULT_ZOOM, OFFICE_MAX_ZOOM, OFFICE_MIN_ZOOM, OFFICE_ZOOM_STEP } from './shell/constants';

const { Content } = Layout;

export type OfficeViewerUri = File | string | (() => Promise<File | Blob | string | Response>);

export type OfficeViewerProps = {
  uri?: OfficeViewerUri;
  defaultFileName?: string;
  defaultPreviewKind?: PreviewKind;
  defaultZoom?: number;
  className?: string;
  style?: CSSProperties;
  onFileParsed?: (parsed: ParsedOfficeFile, file: File) => void;
  onError?: (error: Error, file?: File) => void;
};

const OFFICE_MIME_EXTENSION_MAP: Record<string, string> = {
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
  'application/msword': '.doc',
};

function getFileNameFromUrl(url: string) {
  try {
    const parsedUrl = new URL(url, window.location.href);
    const lastSegment = parsedUrl.pathname.split('/').filter(Boolean).pop();
    return lastSegment ? decodeURIComponent(lastSegment) : undefined;
  } catch {
    const path = url.split(/[?#]/)[0];
    const lastSegment = path.split('/').filter(Boolean).pop();
    return lastSegment ? decodeURIComponent(lastSegment) : undefined;
  }
}

function getFileNameFromContentDisposition(contentDisposition: string | null) {
  if (!contentDisposition) return undefined;

  const encodedMatch = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (encodedMatch?.[1]) return decodeURIComponent(encodedMatch[1]);

  const plainMatch = contentDisposition.match(/filename="?([^";]+)"?/i);
  return plainMatch?.[1] ? decodeURIComponent(plainMatch[1]) : undefined;
}

function getExtensionFromMimeType(mimeType: string) {
  return OFFICE_MIME_EXTENSION_MAP[mimeType.split(';')[0]?.trim().toLowerCase()];
}

function hasFileExtension(fileName: string) {
  return /\.[^./\\]+$/.test(fileName);
}

function ensureSupportedOfficeFile(file: File) {
  if (!isSupportedOfficeFileName(file.name)) {
    throw new Error('暂不支持该文件类型，请选择 PPTX、XLSX、DOCX 或 DOC 文件');
  }
}

function createFileFromBlob(blob: Blob, fileName?: string) {
  if (fileName && hasFileExtension(fileName) && !isSupportedOfficeFileName(fileName)) {
    throw new Error('无法识别 Office 文件类型，请提供 PPTX、XLSX、DOCX 或 DOC 文件');
  }

  const extension = getExtensionFromMimeType(blob.type);
  const inferredFileName =
    fileName && isSupportedOfficeFileName(fileName) ? fileName : extension ? `office-file${extension}` : undefined;

  if (!inferredFileName) {
    throw new Error('无法识别 Office 文件类型，请提供 PPTX、XLSX、DOCX 或 DOC 文件');
  }

  return new File([blob], inferredFileName, { type: blob.type });
}

async function createFileFromResponse(response: Response, fallbackFileName?: string) {
  if (!response.ok) {
    throw new Error(`文件下载失败：${response.status} ${response.statusText}`);
  }

  const blob = await response.blob();
  const fileName =
    getFileNameFromContentDisposition(response.headers.get('Content-Disposition')) || fallbackFileName;
  return createFileFromBlob(blob, fileName);
}

async function downloadOfficeFile(url: string) {
  const urlFileName = getFileNameFromUrl(url);
  if (urlFileName && hasFileExtension(urlFileName) && !isSupportedOfficeFileName(urlFileName)) {
    throw new Error('暂不支持该文件类型，请选择 PPTX、XLSX、DOCX 或 DOC 文件');
  }

  const response = await fetch(url);
  return createFileFromResponse(response, urlFileName);
}

async function normalizeOfficeUri(uri: OfficeViewerUri) {
  const resolvedUri = typeof uri === 'function' ? await uri() : uri;

  if (resolvedUri instanceof File) return resolvedUri;
  if (resolvedUri instanceof Response) return createFileFromResponse(resolvedUri);
  if (resolvedUri instanceof Blob) return createFileFromBlob(resolvedUri);
  if (typeof resolvedUri === 'string') return downloadOfficeFile(resolvedUri);

  throw new Error('uri 必须是 File、URL 字符串，或返回 File/Blob/URL/Response 的异步函数');
}

function OfficeViewerComponent({
  uri,
  defaultFileName = '未加载文件',
  defaultPreviewKind = 'pptx',
  defaultZoom = OFFICE_DEFAULT_ZOOM,
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

  const handleSelectFile = useCallback(
    async (file: File) => {
      setLoading(true);
      setError(undefined);

      try {
        ensureSupportedOfficeFile(file);
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
    if (!uri) return;

    let ignore = false;

    async function loadUri() {
      setLoading(true);
      setError(undefined);

      let file: File | undefined;

      try {
        file = await normalizeOfficeUri(uri);
        if (ignore) return;
        await handleSelectFile(file);
      } catch (nextError) {
        if (ignore) return;
        const normalizedError = nextError instanceof Error ? nextError : new Error('文件加载失败');
        setError(normalizedError.message);
        onError?.(normalizedError, file);
        setLoading(false);
      }
    }

    void loadUri();

    return () => {
      ignore = true;
    };
  }, [handleSelectFile, onError, uri]);

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
        zoom={zoom}
        hasDocument={hasDocument}
        canGoPreviousSlide={canGoPreviousSlide}
        canGoNextSlide={canGoNextSlide}
        onSelectFile={handleSelectFile}
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
