import type { PreviewKind } from '../../preview';
import { deserializeParseError } from '../protocol/errors';
import type {
  MainToWorkerMessage,
  WorkerToMainMessage,
} from '../protocol/messages';
import { OFFICE_PARSER_PROTOCOL_VERSION } from '../protocol/version';
import type { RuntimeSink } from './types';
import { createParseAbortError } from './types';

let taskSequence = 0;

type RuntimeError = Error & {
  code: string;
  recoverable: boolean;
};

function createRuntimeError(
  code: string,
  message: string,
  recoverable: boolean,
): RuntimeError {
  const error = new Error(message) as RuntimeError;
  error.name = 'OfficeWorkerError';
  error.code = code;
  error.recoverable = recoverable;
  return error;
}

/** 判断错误是否发生在文件缓冲区移交 Worker 之前。 */
export function isWorkerStartupError(error: unknown) {
  return (
    error instanceof Error &&
    'code' in error &&
    (error as { code?: unknown }).code === 'WORKER_STARTUP_FAILED'
  );
}

/** 管理单次解析使用的独立 Worker 和跨线程消息。 */
export class WorkerRuntime {
  private worker: Worker | undefined;
  private stopActive: (() => void) | undefined;

  constructor(private readonly workerFactory?: () => Worker) {}

  run(
    file: File,
    kind: 'xls' | 'ppt' | 'doc',
    signal: AbortSignal,
    sink: RuntimeSink,
  ): Promise<void> {
    if (this.worker) {
      return Promise.reject(
        createRuntimeError(
          'WORKER_BUSY',
          '解析 Worker 正在处理其他任务',
          false,
        ),
      );
    }

    return new Promise<void>((resolve, reject) => {
      let worker: Worker;
      try {
        worker = this.workerFactory
          ? this.workerFactory()
          : new Worker(new URL('./worker/entry.js', import.meta.url), {
              type: 'module',
              name: 'office-x-viewer-parser',
            });
      } catch {
        reject(
          createRuntimeError(
            'WORKER_STARTUP_FAILED',
            '无法创建 Office 解析 Worker',
            true,
          ),
        );
        return;
      }

      const taskId = `office-parse-${Date.now()}-${++taskSequence}`;
      let parseStarted = false;
      let settled = false;
      this.worker = worker;

      const cleanup = () => {
        signal.removeEventListener('abort', handleAbort);
        worker.removeEventListener('message', handleMessage);
        worker.removeEventListener('error', handleWorkerError);
        worker.terminate();
        if (this.worker === worker) this.worker = undefined;
        this.stopActive = undefined;
      };
      const fail = (error: unknown) => {
        if (settled) return;
        settled = true;
        cleanup();
        reject(error);
      };
      const finish = () => {
        if (settled) return;
        settled = true;
        cleanup();
        resolve();
      };
      const sendAck = (sequence: number) => {
        const message: MainToWorkerMessage = {
          type: 'parse-ack',
          version: OFFICE_PARSER_PROTOCOL_VERSION,
          taskId,
          sequence,
        };
        worker.postMessage(message);
      };
      const handleAbort = () => {
        if (!settled) {
          const message: MainToWorkerMessage = {
            type: 'parse-cancel',
            version: OFFICE_PARSER_PROTOCOL_VERSION,
            taskId,
          };
          worker.postMessage(message);
        }
        fail(createParseAbortError());
      };
      const handleWorkerError = () => {
        fail(
          createRuntimeError(
            parseStarted ? 'WORKER_RUNTIME_CRASH' : 'WORKER_STARTUP_FAILED',
            parseStarted
              ? 'Office 解析 Worker 运行异常'
              : 'Office 解析 Worker 加载失败',
            !parseStarted,
          ),
        );
      };
      const processMessage = async (message: WorkerToMainMessage) => {
        if (settled) return;
        if (message.version !== OFFICE_PARSER_PROTOCOL_VERSION) {
          fail(
            createRuntimeError(
              'WORKER_STARTUP_FAILED',
              'Office 解析 Worker 协议版本不匹配',
              true,
            ),
          );
          return;
        }
        if (message.type === 'worker-ready') {
          const buffer = await file.arrayBuffer();
          if (settled || signal.aborted) return;
          const startMessage: MainToWorkerMessage = {
            type: 'parse-start',
            version: OFFICE_PARSER_PROTOCOL_VERSION,
            taskId,
            kind,
            fileName: file.name,
            buffer,
          };
          parseStarted = true;
          worker.postMessage(startMessage, [buffer]);
          return;
        }
        if (message.taskId !== taskId) return;
        if (message.type === 'parse-progress') {
          sink.progress(message.progress);
          return;
        }
        if (message.type === 'parse-resource') {
          await sink.resource(message.resource);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-sheet') {
          await sink.sheet(
            message.sheetIndex,
            message.revision,
            message.sheet,
          );
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-presentation-meta') {
          await sink.presentationMetadata(message.metadata);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-slide') {
          await sink.slide(message.slideIndex, message.slide);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-document-meta') {
          await sink.documentMetadata(message.metadata);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-document-blocks') {
          await sink.documentBlocks(message.startIndex, message.blocks);
          sendAck(message.sequence);
          return;
        }
        if (message.type === 'parse-complete') {
          sink.complete(message.warnings);
          finish();
          return;
        }
        if (message.type === 'parse-cancelled') {
          fail(createParseAbortError());
          return;
        }
        if (message.type === 'parse-error') {
          fail(deserializeParseError(message.error));
        }
      };
      const handleMessage = (event: MessageEvent<WorkerToMainMessage>) => {
        void processMessage(event.data).catch(fail);
      };

      this.stopActive = handleAbort;
      signal.addEventListener('abort', handleAbort, { once: true });
      worker.addEventListener('message', handleMessage);
      worker.addEventListener('error', handleWorkerError);
      if (signal.aborted) handleAbort();
    });
  }

  dispose() {
    this.stopActive?.();
    this.worker?.terminate();
    this.worker = undefined;
    this.stopActive = undefined;
  }
}

/** 创建带稳定错误码的 Worker 配置错误。 */
export function createWorkerConfigurationError(
  code: 'WORKER_FORMAT_NOT_READY' | 'WORKER_UNAVAILABLE',
  message: string,
) {
  return createRuntimeError(code, message, false);
}
