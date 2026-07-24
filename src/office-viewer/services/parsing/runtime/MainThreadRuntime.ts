import type { ParsedOfficeFile, PreviewKind } from '../../preview';
import { parseDocCore } from '../../doc/parseDocCore';
import { parsePptCore } from '../../ppt/parsePptCore';
import { parseXlsCore } from '../../xls/parseXlsCore';
import type { RuntimeSink } from './types';
import { createParseAbortError } from './types';

async function parseExistingFormat(
  file: File,
  kind: Exclude<PreviewKind, 'xls' | 'ppt' | 'doc'>,
): Promise<ParsedOfficeFile> {
  if (kind === 'xlsx') {
    const { parseXlsx } = await import('../../xlsx/parseXlsx');
    return { kind, workbook: await parseXlsx(file) };
  }
  if (kind === 'docx') {
    const { parseDocx } = await import('../../docx/parseDocx');
    return { kind, document: await parseDocx(file) };
  }
  const { parsePptx } = await import('../../pptx/parsePptx');
  return { kind, document: await parsePptx(file) };
}

function ensureNotAborted(signal: AbortSignal) {
  if (signal.aborted) throw createParseAbortError();
}

function createDocCheckpoint(signal: AbortSignal, sink: RuntimeSink) {
  let deadline = Date.now() + 12;
  return async (progress?: Parameters<RuntimeSink['progress']>[0]) => {
    ensureNotAborted(signal);
    if (progress) sink.progress(progress);
    if (Date.now() < deadline) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    ensureNotAborted(signal);
    deadline = Date.now() + 12;
  };
}

/** 使用统一事件接口在主线程执行格式解析。 */
export class MainThreadRuntime {
  async run(
    file: File,
    kind: PreviewKind,
    signal: AbortSignal,
    sink: RuntimeSink,
  ) {
    try {
      ensureNotAborted(signal);
      if (kind !== 'xls' && kind !== 'ppt' && kind !== 'doc') {
        sink.progress({
          stage: 'content',
          percent: 0.05,
          message: '正在解析文件',
        });
        const parsed = await parseExistingFormat(file, kind);
        ensureNotAborted(signal);
        await sink.parsed(parsed);
        sink.complete();
        return;
      }

      if (kind === 'doc') {
        sink.progress({
          stage: 'reading',
          percent: 0.01,
          message: '正在读取 DOC/WPS 文件',
        });
        const input = await file.arrayBuffer();
        ensureNotAborted(signal);
        const checkpoint = createDocCheckpoint(signal, sink);
        await parseDocCore(input, {
          fileName: file.name,
          checkpoint,
          output: {
            resource: async (resource) => {
              ensureNotAborted(signal);
              await sink.resource(resource);
            },
            documentMetadata: async (metadata) => {
              ensureNotAborted(signal);
              await sink.documentMetadata(metadata);
            },
            documentBlocks: async (startIndex, blocks) => {
              ensureNotAborted(signal);
              await sink.documentBlocks(startIndex, blocks);
            },
          },
        });
        sink.complete();
        return;
      }

      if (kind === 'ppt') {
        sink.progress({
          stage: 'reading',
          percent: 0.01,
          message: '正在读取 PPT 文件',
        });
        const input = await file.arrayBuffer();
        ensureNotAborted(signal);
        await parsePptCore(input, {
          checkpoint: async (progress) => {
            ensureNotAborted(signal);
            if (progress) sink.progress(progress);
          },
          output: {
            resource: async (resource) => {
              ensureNotAborted(signal);
              await sink.resource(resource);
            },
            presentationMetadata: async (metadata) => {
              ensureNotAborted(signal);
              await sink.presentationMetadata(metadata);
            },
            slide: async (index, slide) => {
              ensureNotAborted(signal);
              await sink.slide(index, slide);
            },
          },
        });
        sink.complete();
        return;
      }

      sink.progress({
        stage: 'reading',
        percent: 0.01,
        message: '正在读取 XLS 文件',
      });
      const input = await file.arrayBuffer();
      ensureNotAborted(signal);
      const result = await parseXlsCore(input, {
        checkpoint: async (progress) => {
          ensureNotAborted(signal);
          if (progress) sink.progress(progress);
        },
        output: {
          resource: async (resource) => {
            ensureNotAborted(signal);
            await sink.resource(resource);
          },
          sheet: async (index, revision, sheet) => {
            ensureNotAborted(signal);
            await sink.sheet(index, revision, sheet);
          },
        },
      });
      sink.complete(result.workbook.warnings);
    } catch (error) {
      sink.error(error);
      throw error;
    }
  }

  dispose() {}
}
