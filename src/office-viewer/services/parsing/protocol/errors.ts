import type { PreviewKind } from '../../preview';
import type { ParseStage } from '../types';

export type SerializedParseError = {
  code: string;
  message: string;
  format?: PreviewKind;
  stage?: ParseStage;
  offset?: number;
  recoverable: boolean;
};

type ErrorWithContext = Error & {
  code?: unknown;
  offset?: unknown;
};

/** 将运行时异常收敛为可安全跨线程传输的错误信息。 */
export function serializeParseError(
  error: unknown,
  context: {
    format?: PreviewKind;
    stage?: ParseStage;
    recoverable?: boolean;
  } = {},
): SerializedParseError {
  const normalized =
    error instanceof Error ? (error as ErrorWithContext) : undefined;
  return {
    code:
      typeof normalized?.code === 'string'
        ? normalized.code
        : 'WORKER_PARSE_FAILED',
    message: normalized?.message ?? '文件解析失败',
    format: context.format,
    stage: context.stage,
    offset:
      typeof normalized?.offset === 'number' ? normalized.offset : undefined,
    recoverable: context.recoverable ?? false,
  };
}

/** 将跨线程错误恢复为 Error，并保留稳定错误码供调用方判断。 */
export function deserializeParseError(source: SerializedParseError): Error {
  const error = new Error(source.message) as Error & {
    code: string;
    format?: PreviewKind;
    stage?: ParseStage;
    offset?: number;
    recoverable: boolean;
  };
  error.name = 'OfficeParseError';
  error.code = source.code;
  error.format = source.format;
  error.stage = source.stage;
  error.offset = source.offset;
  error.recoverable = source.recoverable;
  return error;
}
