export type CfbParseErrorCode =
  | 'INVALID_SIGNATURE'
  | 'INVALID_HEADER'
  | 'SECTOR_OUT_OF_RANGE'
  | 'CHAIN_CYCLE'
  | 'CHAIN_TRUNCATED'
  | 'DIRECTORY_CORRUPTED';

/** 表示可安全展示的 CFB 结构错误，不包含本地文件路径。 */
export class CfbParseError extends Error {
  readonly code: CfbParseErrorCode;
  readonly sector?: number;
  readonly directoryId?: number;

  constructor(
    code: CfbParseErrorCode,
    message: string,
    context: { sector?: number; directoryId?: number } = {},
  ) {
    super(message);
    this.name = 'CfbParseError';
    this.code = code;
    this.sector = context.sector;
    this.directoryId = context.directoryId;
  }
}
