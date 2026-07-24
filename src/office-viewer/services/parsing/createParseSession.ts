import type { ParsedOfficeFile } from '../preview';
import {
  DocDocumentAssembler,
  PptDocumentAssembler,
  XlsDocumentAssembler,
} from './assembly/DocumentAssembler';
import { ResourceRegistry } from './assembly/ResourceRegistry';
import { detectPreviewKind } from './detectPreviewKind';
import type { OfficeViewerParseSession } from './internalTypes';
import { MainThreadRuntime } from './runtime/MainThreadRuntime';
import { createRuntime } from './runtime/createRuntime';
import type { RuntimeSink } from './runtime/types';
import {
  isWorkerStartupError,
  WorkerRuntime,
} from './runtime/WorkerRuntime';
import type {
  OfficeParseOptions,
  OfficeParseSession,
  OfficeParseSessionStatus,
  ParseProgress,
} from './types';

function createParseSession(
  file: File,
  options: OfficeParseOptions,
  enablePartial: boolean,
): OfficeViewerParseSession {
  const kind = detectPreviewKind(file.name);
  const controller = new AbortController();
  const listeners = new Set<(progress: ParseProgress) => void>();
  const partialListeners = new Set<(parsed: ParsedOfficeFile) => void>();
  const assembler =
    kind === 'xls'
      ? new XlsDocumentAssembler(new ResourceRegistry())
      : undefined;
  const presentationAssembler =
    kind === 'ppt'
      ? new PptDocumentAssembler(new ResourceRegistry())
      : undefined;
  const documentAssembler =
    kind === 'doc'
      ? new DocDocumentAssembler(new ResourceRegistry())
      : undefined;
  let runtime: MainThreadRuntime | WorkerRuntime | undefined;
  let status: OfficeParseSessionStatus = 'starting';
  let parsedResult: ParsedOfficeFile | undefined;
  let partialResult: ParsedOfficeFile | undefined;

  const emitProgress = (progress: ParseProgress) => {
    listeners.forEach((listener) => {
      try {
        listener(progress);
      } catch (listenerError) {
        // 调用方进度回调不能破坏解析任务本身。
        void listenerError;
      }
    });
  };

  const emitPartial = (parsed: ParsedOfficeFile) => {
    if (!enablePartial) return;
    partialListeners.forEach((listener) => {
      try {
        listener(parsed);
      } catch (listenerError) {
        // 调用方的渐进渲染异常不能中断底层文件解析。
        void listenerError;
      }
    });
  };

  const sink: RuntimeSink = {
    progress: emitProgress,
    resource: async (resource) => {
      const target = assembler ?? presentationAssembler ?? documentAssembler;
      if (!target) throw new Error('当前格式会话收到了资源分块');
      await target.addResource(resource);
    },
    sheet: async (index, revision, sheet) => {
      if (!assembler) throw new Error('非 XLS 会话收到了工作表分块');
      assembler.addSheet(index, revision, sheet);
      if (enablePartial && assembler.hasRenderableContent()) {
        emitPartial({ kind: 'xls', workbook: assembler.snapshot() });
      }
    },
    presentationMetadata: async (metadata) => {
      if (!presentationAssembler) {
        throw new Error('非 PPT 会话收到了演示文稿元数据');
      }
      presentationAssembler.setMetadata(metadata);
      if (presentationAssembler.hasRenderableContent()) {
        emitPartial({
          kind: 'ppt',
          document: presentationAssembler.snapshot(),
        });
      }
    },
    slide: async (index, slide) => {
      if (!presentationAssembler) {
        throw new Error('非 PPT 会话收到了幻灯片分块');
      }
      presentationAssembler.addSlide(index, slide);
      if (presentationAssembler.hasRenderableContent()) {
        emitPartial({
          kind: 'ppt',
          document: presentationAssembler.snapshot(),
        });
      }
    },
    documentMetadata: async (metadata) => {
      if (!documentAssembler) {
        throw new Error('非 DOC 会话收到了文档元数据');
      }
      documentAssembler.setMetadata(metadata);
      if (documentAssembler.hasRenderableContent()) {
        emitPartial({
          kind: 'doc',
          document: documentAssembler.snapshot(),
        });
      }
    },
    documentBlocks: async (startIndex, blocks) => {
      if (!documentAssembler) {
        throw new Error('非 DOC 会话收到了正文分块');
      }
      documentAssembler.addBlocks(startIndex, blocks);
      if (documentAssembler.hasRenderableContent()) {
        emitPartial({
          kind: 'doc',
          document: documentAssembler.snapshot(),
        });
      }
    },
    parsed: async (parsed) => {
      parsedResult = parsed;
    },
    complete: (warnings) => {
      if (assembler) {
        assembler.setWarnings(warnings);
        parsedResult = { kind: 'xls', workbook: assembler.complete() };
      } else if (presentationAssembler) {
        parsedResult = {
          kind: 'ppt',
          document: presentationAssembler.complete(),
        };
      } else if (documentAssembler) {
        parsedResult = {
          kind: 'doc',
          document: documentAssembler.complete(),
        };
      }
    },
    error: () => undefined,
  };

  const run = async () => {
    status = 'running';
    const workerMode = options.worker ?? 'auto';
    runtime = createRuntime(workerMode, kind, options.workerFactory);
    try {
      if (runtime instanceof WorkerRuntime) {
        try {
          if (kind !== 'xls' && kind !== 'ppt' && kind !== 'doc') {
            throw new Error('当前格式尚未启用 Worker');
          }
          await runtime.run(file, kind, controller.signal, sink);
        } catch (error) {
          if (workerMode !== 'auto' || !isWorkerStartupError(error)) throw error;
          runtime.dispose();
          runtime = new MainThreadRuntime();
          await runtime.run(file, kind, controller.signal, sink);
        }
      } else {
        await runtime.run(file, kind, controller.signal, sink);
      }
      if (!parsedResult) throw new Error('解析运行时未返回文档结果');
      status = 'completed';
      return parsedResult;
    } catch (error) {
      const cancelled =
        error instanceof Error && error.name === 'AbortError';
      if (cancelled) {
        status = 'cancelled';
      } else {
        status = 'failed';
        if (enablePartial && assembler?.hasRenderableContent()) {
          partialResult = {
            kind: 'xls',
            workbook: assembler.completePartial(),
          };
        } else if (
          enablePartial &&
          presentationAssembler?.hasRenderableContent()
        ) {
          partialResult = {
            kind: 'ppt',
            document: presentationAssembler.completePartial(),
          };
        } else if (
          enablePartial &&
          documentAssembler?.hasRenderableContent()
        ) {
          partialResult = {
            kind: 'doc',
            document: documentAssembler.completePartial(),
          };
        }
      }
      if (partialResult?.kind !== 'xls') assembler?.dispose();
      if (partialResult?.kind !== 'ppt') presentationAssembler?.dispose();
      if (partialResult?.kind !== 'doc') documentAssembler?.dispose();
      throw error;
    } finally {
      runtime?.dispose();
      runtime = undefined;
    }
  };

  const result = run();
  return {
    result,
    get status() {
      return status;
    },
    get partialResult() {
      return partialResult;
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    subscribePartial(listener) {
      partialListeners.add(listener);
      return () => partialListeners.delete(listener);
    },
    cancel() {
      if (
        status === 'completed' ||
        status === 'cancelled' ||
        status === 'failed'
      ) {
        return;
      }
      controller.abort();
      runtime?.dispose();
    },
    dispose() {
      if (status === 'starting' || status === 'running') {
        controller.abort();
        runtime?.dispose();
      }
      listeners.clear();
      partialListeners.clear();
      if (partialResult?.kind !== 'xls') assembler?.dispose();
      if (partialResult?.kind !== 'ppt') presentationAssembler?.dispose();
      if (partialResult?.kind !== 'doc') documentAssembler?.dispose();
    },
  };
}

/** 创建单文件解析会话，统一管理运行时、进度、取消和结果资源。 */
export function createOfficeParseSession(
  file: File,
  options: OfficeParseOptions = {},
): OfficeParseSession<ParsedOfficeFile> {
  return createParseSession(file, options, false);
}

/** 创建仅供 OfficeViewer 使用的渐进解析会话。 */
export function createOfficeViewerParseSession(
  file: File,
  options: OfficeParseOptions = {},
): OfficeViewerParseSession {
  return createParseSession(file, options, true);
}
