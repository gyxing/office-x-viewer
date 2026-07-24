export type XlsParseErrorCode =
  | 'INVALID_CFB'
  | 'UNSUPPORTED_BIFF_VERSION'
  | 'ENCRYPTED_FILE'
  | 'MISSING_WORKBOOK_STREAM'
  | 'TRUNCATED_RECORD'
  | 'CORRUPTED_SECTOR_CHAIN'
  | 'INVALID_RECORD_DATA';

/** XLS 解析错误，保留 BIFF 记录上下文供界面展示和排查。 */
export class XlsParseError extends Error {
  readonly code: XlsParseErrorCode;
  readonly offset?: number;
  readonly recordId?: number;

  constructor(
    code: XlsParseErrorCode,
    message: string,
    context: { offset?: number; recordId?: number } = {},
  ) {
    super(message);
    this.name = 'XlsParseError';
    this.code = code;
    this.offset = context.offset;
    this.recordId = context.recordId;
  }
}
