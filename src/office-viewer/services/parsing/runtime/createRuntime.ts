import type { PreviewKind } from '../../preview';
import type { WorkerMode } from '../types';
import { MainThreadRuntime } from './MainThreadRuntime';
import {
  createWorkerConfigurationError,
  WorkerRuntime,
} from './WorkerRuntime';

/** 根据运行模式和格式能力选择解析运行时。 */
export function createRuntime(
  mode: WorkerMode,
  kind: PreviewKind,
  workerFactory?: () => Worker,
) {
  if (mode === 'never') return new MainThreadRuntime();
  if (kind !== 'xls' && kind !== 'ppt' && kind !== 'doc') {
    if (mode === 'always') {
      throw createWorkerConfigurationError(
        'WORKER_FORMAT_NOT_READY',
        `${kind.toUpperCase()} 尚未完成 Worker 迁移`,
      );
    }
    return new MainThreadRuntime();
  }
  if (typeof Worker === 'undefined' && !workerFactory) {
    if (mode === 'always') {
      throw createWorkerConfigurationError(
        'WORKER_UNAVAILABLE',
        '当前环境不支持 Web Worker',
      );
    }
    return new MainThreadRuntime();
  }
  return new WorkerRuntime(workerFactory);
}
