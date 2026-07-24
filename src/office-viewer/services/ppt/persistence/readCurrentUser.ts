import {
  PPT_CURRENT_USER_ENCRYPTED_TOKEN,
  PPT_CURRENT_USER_UNENCRYPTED_TOKEN,
  PPT_RECORD,
} from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { readPptByteString, readPptUnicodeString } from '../binary/readStrings';
import { PptParseError } from '../errors';
import type { PptParseContext } from '../types';

export type PptCurrentUser = {
  offsetToCurrentEdit: number;
  userName?: string;
};

/** 读取 Current User 流，并定位最新一次保存对应的 UserEditAtom。 */
export function readPptCurrentUser(
  bytes: Uint8Array,
  context: PptParseContext,
): PptCurrentUser {
  const record = new PptRecordReader(bytes).readRecord();
  if (!record || record.type !== PPT_RECORD.CURRENT_USER_ATOM) {
    throw new PptParseError(
      'PPT_EDIT_CHAIN_INVALID',
      'Current User 流缺少有效的 CurrentUserAtom',
    );
  }
  if (record.length < 20) {
    throw new PptParseError(
      'PPT_EDIT_CHAIN_INVALID',
      'CurrentUserAtom 长度不足',
      { offset: record.offset, recordType: record.type },
    );
  }

  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  const headerToken = view.getUint32(4, true);
  if (headerToken === PPT_CURRENT_USER_ENCRYPTED_TOKEN) {
    throw new PptParseError('PPT_ENCRYPTED', '暂不支持加密的 PPT 文件');
  }
  if (headerToken !== PPT_CURRENT_USER_UNENCRYPTED_TOKEN) {
    context.warnings.push({
      code: 'PPT_CURRENT_USER_TOKEN_UNKNOWN',
      message: 'Current User 流使用了未知标记，已按未加密文件继续解析',
      offset: record.offset + 12,
    });
  }

  const offsetToCurrentEdit = view.getUint32(8, true);
  const ansiNameLength = view.getUint16(12, true);
  const ansiNameOffset = 20;
  const ansiNameEnd = Math.min(record.data.length, ansiNameOffset + ansiNameLength);
  const ansiName =
    ansiNameEnd > ansiNameOffset
      ? readPptByteString(
          record.data.subarray(ansiNameOffset, ansiNameEnd),
          1252,
          context,
        )
      : undefined;
  const unicodeNameOffset = ansiNameEnd + 4;
  const unicodeNameLength = ansiNameLength * 2;
  const unicodeName =
    unicodeNameOffset + unicodeNameLength <= record.data.length
      ? readPptUnicodeString(
          record.data.subarray(
            unicodeNameOffset,
            unicodeNameOffset + unicodeNameLength,
          ),
        )
      : undefined;

  return {
    offsetToCurrentEdit,
    userName: unicodeName || ansiName,
  };
}
