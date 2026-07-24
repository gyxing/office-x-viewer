// OfficeViewer 是组件库对外主入口，负责文件上传解析、格式状态和全局工具栏交互。
import { Layout } from 'antd';
import type { CSSProperties } from 'react';
import React, {
  memo,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import './index.less';
import { disposeDocDocument, type DocDocument } from './services/doc/types';
import type { DocxDocument } from './services/docx/types';
import {
  type OfficeParseOptions,
  type ParseProgress,
} from './services/parsing';
import { createOfficeViewerParseSession } from './services/parsing/createParseSession';
import type { OfficeViewerParseSession } from './services/parsing/internalTypes';
import type { PptxDocument } from './services/pptx/types';
import { disposePresentationDocument } from './services/presentation/dispose';
import {
  detectPreviewKind,
  isPresentationPreviewKind,
  isSpreadsheetPreviewKind,
  isSupportedOfficeFileName,
  type ParsedOfficeFile,
  type PreviewKind,
} from './services/preview';
import {
  disposeSpreadsheetWorkbook,
  type SpreadsheetWorkbook,
} from './services/spreadsheet/types';
import { OfficeParseStatus } from './shell/ParseStatus';
import { OfficePreviewStage } from './shell/PreviewStage';
import { OfficeToolbar } from './shell/Toolbar';
import {
  OFFICE_DEFAULT_ZOOM,
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_STEP,
} from './shell/constants';

const { Content } = Layout;

type PendingPartialResult = {
  loadGeneration: number;
  parsed: ParsedOfficeFile;
};

export type OfficeViewerUri =
  | File
  | string
  | (() => Promise<File | Blob | string | Response>);

export type OfficeViewerProps = {
  uri?: OfficeViewerUri;
  defaultFileName?: string;
  defaultPreviewKind?: PreviewKind;
  defaultZoom?: number;
  className?: string;
  height?: CSSProperties['height'];
  style?: CSSProperties;
  onFileParsed?: (parsed: ParsedOfficeFile, file: File) => void;
  onError?: (error: Error, file?: File) => void;
  parseOptions?: OfficeParseOptions;
  onParseProgress?: (progress: ParseProgress) => void;
};

const OFFICE_MIME_EXTENSION_MAP: Record<string, string> = {
  'application/vnd.openxmlformats-officedocument.presentationml.presentation':
    '.pptx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
  'application/vnd.ms-excel': '.xls',
  'application/vnd.ms-powerpoint': '.ppt',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
    '.docx',
  'application/msword': '.doc',
  'application/wps-office.wps': '.wps',
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
  return OFFICE_MIME_EXTENSION_MAP[
    mimeType.split(';')[0]?.trim().toLowerCase()
  ];
}

function hasFileExtension(fileName: string) {
  return /\.[^./\\]+$/.test(fileName);
}

function ensureSupportedOfficeFile(file: File) {
  if (!isSupportedOfficeFileName(file.name)) {
    throw new Error(
      '暂不支持该文件类型，请选择 PPTX、PPT、XLSX、XLS、DOCX、DOC 或 WPS 文件',
    );
  }
}

function createFileFromBlob(blob: Blob, fileName?: string) {
  if (
    fileName &&
    hasFileExtension(fileName) &&
    !isSupportedOfficeFileName(fileName)
  ) {
    throw new Error(
      '无法识别 Office 文件类型，请提供 PPTX、PPT、XLSX、XLS、DOCX、DOC 或 WPS 文件',
    );
  }

  const extension = getExtensionFromMimeType(blob.type);
  const inferredFileName =
    fileName && isSupportedOfficeFileName(fileName)
      ? fileName
      : extension
      ? `office-file${extension}`
      : undefined;

  if (!inferredFileName) {
    throw new Error(
      '无法识别 Office 文件类型，请提供 PPTX、PPT、XLSX、XLS、DOCX、DOC 或 WPS 文件',
    );
  }

  return new File([blob], inferredFileName, { type: blob.type });
}

async function createFileFromResponse(
  response: Response,
  fallbackFileName?: string,
) {
  if (!response.ok) {
    throw new Error(`文件下载失败：${response.status} ${response.statusText}`);
  }

  const blob = await response.blob();
  const fileName =
    getFileNameFromContentDisposition(
      response.headers.get('Content-Disposition'),
    ) || fallbackFileName;
  return createFileFromBlob(blob, fileName);
}

async function downloadOfficeFile(url: string, signal?: AbortSignal) {
  const urlFileName = getFileNameFromUrl(url);
  if (
    urlFileName &&
    hasFileExtension(urlFileName) &&
    !isSupportedOfficeFileName(urlFileName)
  ) {
    throw new Error(
      '暂不支持该文件类型，请选择 PPTX、PPT、XLSX、XLS、DOCX、DOC 或 WPS 文件',
    );
  }

  const response = await fetch(url, { signal });
  return createFileFromResponse(response, urlFileName);
}

async function normalizeOfficeUri(uri: OfficeViewerUri, signal?: AbortSignal) {
  const resolvedUri = typeof uri === 'function' ? await uri() : uri;

  if (resolvedUri instanceof File) return resolvedUri;
  if (resolvedUri instanceof Response)
    return createFileFromResponse(resolvedUri);
  if (resolvedUri instanceof Blob) return createFileFromBlob(resolvedUri);
  if (typeof resolvedUri === 'string')
    return downloadOfficeFile(resolvedUri, signal);

  throw new Error(
    'uri 必须是 File、URL 字符串，或返回 File/Blob/URL/Response 的异步函数',
  );
}

/** 释放一个已经取得 Blob URL 所有权的解析结果。 */
function disposeParsedOfficeFile(parsed: ParsedOfficeFile) {
  if (isSpreadsheetPreviewKind(parsed.kind)) {
    disposeSpreadsheetWorkbook(parsed.workbook);
    return;
  }
  if (isPresentationPreviewKind(parsed.kind)) {
    disposePresentationDocument(parsed.document);
    return;
  }
  if (parsed.kind === 'doc') disposeDocDocument(parsed.document);
}

function OfficeViewerComponent({
  uri,
  defaultFileName = '未加载文件',
  defaultPreviewKind = 'pptx',
  defaultZoom = OFFICE_DEFAULT_ZOOM,
  className,
  height,
  style,
  onFileParsed,
  onError,
  parseOptions,
  onParseProgress,
}: OfficeViewerProps) {
  // OfficeViewer 是公共组件入口，集中管理“文件状态”和“格式私有状态”，避免使用者再组合多个子组件。
  const [fileName, setFileName] = useState(defaultFileName);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>();
  const [parseProgress, setParseProgress] = useState<ParseProgress>();
  const [partialWarning, setPartialWarning] = useState<string>();
  const [previewKind, setPreviewKind] =
    useState<PreviewKind>(defaultPreviewKind);
  const [pptxDocument, setPptxDocument] = useState<PptxDocument>();
  const presentationDocumentRef = useRef<PptxDocument>();
  const [spreadsheetWorkbook, setSpreadsheetWorkbook] =
    useState<SpreadsheetWorkbook>();
  const spreadsheetWorkbookRef = useRef<SpreadsheetWorkbook>();
  const [docxDocument, setDocxDocument] = useState<DocxDocument>();
  const [docDocument, setDocDocument] = useState<DocDocument>();
  const docDocumentRef = useRef<DocDocument>();
  const [activeIndex, setActiveIndex] = useState(0);
  const [activeSheetId, setActiveSheetId] = useState<string>();
  const [zoom, setZoom] = useState(defaultZoom);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const viewerRef = useRef<HTMLDivElement | null>(null);
  const loadGenerationRef = useRef(0);
  const requestControllerRef = useRef<AbortController>();
  const parseSessionRef = useRef<OfficeViewerParseSession>();
  const pendingPartialRef = useRef<PendingPartialResult>();
  const partialFrameRef = useRef<number>();
  const defaultZoomRef = useRef(defaultZoom);
  const onFileParsedRef = useRef(onFileParsed);
  const onErrorRef = useRef(onError);
  const parseOptionsRef = useRef(parseOptions);
  const onParseProgressRef = useRef(onParseProgress);

  defaultZoomRef.current = defaultZoom;
  onFileParsedRef.current = onFileParsed;
  onErrorRef.current = onError;
  parseOptionsRef.current = parseOptions;
  onParseProgressRef.current = onParseProgress;

  const cancelPartialFrame = useCallback(() => {
    if (
      partialFrameRef.current !== undefined &&
      typeof window !== 'undefined'
    ) {
      window.cancelAnimationFrame(partialFrameRef.current);
    }
    partialFrameRef.current = undefined;
    pendingPartialRef.current = undefined;
  }, []);

  const clearPreviewDocuments = useCallback(() => {
    // 先移除 React 模型，再释放 refs 中唯一拥有资源的完整或失败冻结模型。
    setPptxDocument(undefined);
    setSpreadsheetWorkbook(undefined);
    setDocxDocument(undefined);
    setDocDocument(undefined);
    setActiveIndex(0);
    setActiveSheetId(undefined);

    const spreadsheet = spreadsheetWorkbookRef.current;
    spreadsheetWorkbookRef.current = undefined;
    disposeSpreadsheetWorkbook(spreadsheet);
    const presentation = presentationDocumentRef.current;
    presentationDocumentRef.current = undefined;
    disposePresentationDocument(presentation);
    const documentModel = docDocumentRef.current;
    docDocumentRef.current = undefined;
    disposeDocDocument(documentModel);
  }, []);

  const installPartialSnapshot = useCallback((parsed: ParsedOfficeFile) => {
    if (parsed.kind === 'xls') {
      setSpreadsheetWorkbook(parsed.workbook);
      setActiveSheetId((current) =>
        current && parsed.workbook.sheets.some((sheet) => sheet.id === current)
          ? current
          : parsed.workbook.sheets[0]?.id,
      );
      return;
    }
    if (parsed.kind === 'ppt') {
      setPptxDocument(parsed.document);
      setActiveIndex((current) =>
        Math.min(current, Math.max(0, parsed.document.slides.length - 1)),
      );
      return;
    }
    if (parsed.kind === 'doc') {
      setDocDocument(parsed.document);
    }
  }, []);

  const schedulePartialSnapshot = useCallback(
    (parsed: ParsedOfficeFile, loadGeneration: number) => {
      pendingPartialRef.current = { parsed, loadGeneration };
      if (partialFrameRef.current !== undefined) return;
      if (typeof window === 'undefined') {
        if (loadGeneration === loadGenerationRef.current) {
          installPartialSnapshot(parsed);
        }
        pendingPartialRef.current = undefined;
        return;
      }
      partialFrameRef.current = window.requestAnimationFrame(() => {
        partialFrameRef.current = undefined;
        const pending = pendingPartialRef.current;
        pendingPartialRef.current = undefined;
        if (pending && pending.loadGeneration === loadGenerationRef.current) {
          installPartialSnapshot(pending.parsed);
        }
      });
    },
    [installPartialSnapshot],
  );

  const loadFile = useCallback(
    async (file: File, loadGeneration: number) => {
      parseSessionRef.current?.cancel();
      parseSessionRef.current?.dispose();
      parseSessionRef.current = undefined;
      cancelPartialFrame();
      clearPreviewDocuments();
      setLoading(true);
      setError(undefined);
      setParseProgress(undefined);
      setPartialWarning(undefined);
      let retainedPartial = false;

      try {
        ensureSupportedOfficeFile(file);
        if (loadGeneration !== loadGenerationRef.current) return;

        // 上传新文件时同步重置所有格式相关状态，防止上一份文档的页码/缩放/工作表残留到新文档。
        const fileKind = detectPreviewKind(file.name);
        setPreviewKind(fileKind);
        setFileName(file.name);
        setActiveIndex(0);
        setZoom(defaultZoomRef.current);

        const parseSession = createOfficeViewerParseSession(
          file,
          parseOptionsRef.current,
        );
        parseSessionRef.current = parseSession;
        const unsubscribeProgress = parseSession.subscribe((progress) => {
          if (loadGeneration !== loadGenerationRef.current) return;
          setParseProgress(progress);
          onParseProgressRef.current?.(progress);
        });
        const unsubscribePartial = parseSession.subscribePartial((partial) => {
          if (loadGeneration !== loadGenerationRef.current) return;
          schedulePartialSnapshot(partial, loadGeneration);
        });
        let parsed: ParsedOfficeFile;
        try {
          parsed = await parseSession.result;
        } catch (nextError) {
          cancelPartialFrame();
          const partial = parseSession.partialResult;
          if (partial) {
            if (loadGeneration !== loadGenerationRef.current) {
              disposeParsedOfficeFile(partial);
            } else if (partial.kind === 'xls') {
              const normalizedError =
                nextError instanceof Error
                  ? nextError
                  : new Error('文件解析失败');
              spreadsheetWorkbookRef.current = partial.workbook;
              installPartialSnapshot(partial);
              setPartialWarning(normalizedError.message);
              setError(undefined);
              retainedPartial = true;
            } else if (partial.kind === 'ppt') {
              const normalizedError =
                nextError instanceof Error
                  ? nextError
                  : new Error('文件解析失败');
              presentationDocumentRef.current = partial.document;
              installPartialSnapshot(partial);
              setPartialWarning(normalizedError.message);
              setError(undefined);
              retainedPartial = true;
            } else if (partial.kind === 'doc') {
              const normalizedError =
                nextError instanceof Error
                  ? nextError
                  : new Error('文件解析失败');
              docDocumentRef.current = partial.document;
              installPartialSnapshot(partial);
              setPartialWarning(normalizedError.message);
              setError(undefined);
              retainedPartial = true;
            } else {
              disposeParsedOfficeFile(partial);
            }
          }
          throw nextError;
        } finally {
          unsubscribeProgress();
          unsubscribePartial();
          if (parseSessionRef.current === parseSession) {
            parseSession.dispose();
            parseSessionRef.current = undefined;
          }
        }
        if (loadGeneration !== loadGenerationRef.current) {
          disposeParsedOfficeFile(parsed);
          return;
        }

        cancelPartialFrame();
        setParseProgress(undefined);
        setPartialWarning(undefined);
        const nextPresentationDocument = isPresentationPreviewKind(parsed.kind)
          ? parsed.document
          : undefined;
        disposePresentationDocument(presentationDocumentRef.current);
        presentationDocumentRef.current = nextPresentationDocument;
        setPptxDocument(nextPresentationDocument);
        const nextSpreadsheetWorkbook = isSpreadsheetPreviewKind(parsed.kind)
          ? parsed.workbook
          : undefined;
        disposeSpreadsheetWorkbook(spreadsheetWorkbookRef.current);
        spreadsheetWorkbookRef.current = nextSpreadsheetWorkbook;
        setSpreadsheetWorkbook(nextSpreadsheetWorkbook);
        setDocxDocument(parsed.kind === 'docx' ? parsed.document : undefined);
        const nextDocDocument =
          parsed.kind === 'doc' ? parsed.document : undefined;
        disposeDocDocument(docDocumentRef.current);
        docDocumentRef.current = nextDocDocument;
        setDocDocument(nextDocDocument);
        setActiveSheetId((current) => {
          if (!isSpreadsheetPreviewKind(parsed.kind)) return undefined;
          return current &&
            parsed.workbook.sheets.some((sheet) => sheet.id === current)
            ? current
            : parsed.workbook.sheets[0]?.id;
        });
        onFileParsedRef.current?.(parsed, file);
      } catch (nextError) {
        if (loadGeneration !== loadGenerationRef.current) return;

        // 对外回调始终给 Error 实例，组件内部只保存可展示的 message。
        const normalizedError =
          nextError instanceof Error ? nextError : new Error('文件解析失败');
        setParseProgress(undefined);
        if (!retainedPartial) setError(normalizedError.message);
        onErrorRef.current?.(normalizedError, file);
      } finally {
        if (loadGeneration === loadGenerationRef.current) {
          setLoading(false);
        }
      }
    },
    [
      cancelPartialFrame,
      clearPreviewDocuments,
      installPartialSnapshot,
      schedulePartialSnapshot,
    ],
  );

  const handleSelectFile = useCallback(
    async (file: File) => {
      requestControllerRef.current?.abort();
      requestControllerRef.current = undefined;
      const loadGeneration = ++loadGenerationRef.current;
      await loadFile(file, loadGeneration);
    },
    [loadFile],
  );

  useEffect(() => {
    if (!uri) return;

    parseSessionRef.current?.cancel();
    parseSessionRef.current?.dispose();
    parseSessionRef.current = undefined;
    cancelPartialFrame();
    clearPreviewDocuments();
    setParseProgress(undefined);
    setPartialWarning(undefined);
    // 固化本次 effect 的文件来源，避免异步闭包丢失类型收窄。
    const uriToLoad = uri;
    const loadGeneration = ++loadGenerationRef.current;
    const requestController =
      typeof AbortController === 'undefined'
        ? undefined
        : new AbortController();
    requestControllerRef.current?.abort();
    requestControllerRef.current = requestController;

    async function loadUri() {
      setLoading(true);
      setError(undefined);

      let file: File | undefined;

      try {
        file = await normalizeOfficeUri(uriToLoad, requestController?.signal);
        if (loadGeneration !== loadGenerationRef.current) return;
        await loadFile(file, loadGeneration);
      } catch (nextError) {
        if (
          loadGeneration !== loadGenerationRef.current ||
          requestController?.signal.aborted
        )
          return;

        const normalizedError =
          nextError instanceof Error ? nextError : new Error('文件加载失败');
        setError(normalizedError.message);
        onErrorRef.current?.(normalizedError, file);
        setLoading(false);
      } finally {
        if (requestControllerRef.current === requestController) {
          requestControllerRef.current = undefined;
        }
      }
    }

    void loadUri();

    return () => {
      requestController?.abort();
    };
  }, [cancelPartialFrame, clearPreviewDocuments, loadFile, uri]);

  useEffect(
    () => () => {
      // 组件卸载时让所有不可取消的本地解析任务失效，避免异步结果继续写入状态。
      loadGenerationRef.current += 1;
      requestControllerRef.current?.abort();
      parseSessionRef.current?.cancel();
      parseSessionRef.current?.dispose();
      parseSessionRef.current = undefined;
      cancelPartialFrame();
      disposeSpreadsheetWorkbook(spreadsheetWorkbookRef.current);
      spreadsheetWorkbookRef.current = undefined;
      disposePresentationDocument(presentationDocumentRef.current);
      presentationDocumentRef.current = undefined;
      disposeDocDocument(docDocumentRef.current);
      docDocumentRef.current = undefined;
    },
    [cancelPartialFrame],
  );

  const hasDocument = useMemo(
    () =>
      // 工具栏的翻页/缩放按钮只依赖“当前格式是否有可渲染内容”，不要耦合到具体 viewer 实现。
      isPresentationPreviewKind(previewKind)
        ? Boolean(pptxDocument?.slides.length)
        : isSpreadsheetPreviewKind(previewKind)
        ? Boolean(spreadsheetWorkbook?.sheets.length)
        : previewKind === 'docx'
        ? Boolean(docxDocument?.blocks.length)
        : Boolean(docDocument?.paragraphs.length),
    [docDocument, docxDocument, pptxDocument, previewKind, spreadsheetWorkbook],
  );

  const canGoPreviousSlide =
    isPresentationPreviewKind(previewKind) &&
    Boolean(pptxDocument?.slides.length) &&
    activeIndex > 0;
  const canGoNextSlide =
    isPresentationPreviewKind(previewKind) &&
    Boolean(pptxDocument?.slides.length) &&
    activeIndex < (pptxDocument?.slides.length ?? 1) - 1;

  const handlePreviousSlide = useCallback(() => {
    setActiveIndex((value) => Math.max(value - 1, 0));
  }, []);

  const handleNextSlide = useCallback(() => {
    setActiveIndex((value) =>
      Math.min(value + 1, (pptxDocument?.slides.length ?? 1) - 1),
    );
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

  const fullscreenSupported =
    typeof document !== 'undefined' &&
    typeof document.documentElement.requestFullscreen === 'function';

  useEffect(() => {
    if (typeof document === 'undefined') return;

    // 浏览器和 ESC 键都可能改变全屏状态，因此以 fullscreenchange 作为唯一状态来源。
    const handleFullscreenChange = () => {
      setIsFullscreen(document.fullscreenElement === viewerRef.current);
    };

    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => {
      document.removeEventListener('fullscreenchange', handleFullscreenChange);
    };
  }, []);

  const handleFullscreen = useCallback(async () => {
    const viewer = viewerRef.current;
    if (
      !viewer ||
      typeof document === 'undefined' ||
      typeof viewer.requestFullscreen !== 'function'
    )
      return;

    try {
      if (document.fullscreenElement === viewer) {
        await document.exitFullscreen();
      } else {
        await viewer.requestFullscreen();
      }
    } catch (nextError) {
      const message =
        nextError instanceof Error ? nextError.message : '浏览器拒绝了全屏请求';
      onErrorRef.current?.(new Error(`全屏操作失败：${message}`));
    }
  }, []);

  // 专用 height 配置优先于 style.height，避免两个入口同时传值时结果不确定。
  const viewerStyle = height === undefined ? style : { ...style, height };

  return (
    <div
      ref={viewerRef}
      className={['oxv-office-viewer', className].filter(Boolean).join(' ')}
      style={viewerStyle}
    >
      <Layout className="oxv-office-viewer__layout">
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
          isFullscreen={isFullscreen}
          fullscreenSupported={fullscreenSupported}
          onFullscreen={handleFullscreen}
        />
        <Content className="oxv-office-viewer__content">
          <OfficePreviewStage
            loading={loading}
            loadingTip={parseProgress?.message}
            hasRenderableContent={hasDocument}
            error={error}
            previewKind={previewKind}
            pptxDocument={pptxDocument}
            spreadsheetWorkbook={spreadsheetWorkbook}
            docxDocument={docxDocument}
            docDocument={docDocument}
            activeIndex={activeIndex}
            activeSheetId={activeSheetId}
            zoom={zoom}
            onSelectSlide={setActiveIndex}
            onSelectSheet={setActiveSheetId}
          />
          <OfficeParseStatus
            progress={loading && hasDocument ? parseProgress : undefined}
            warning={partialWarning}
          />
        </Content>
      </Layout>
    </div>
  );
}

export const OfficeViewer = memo(OfficeViewerComponent);
