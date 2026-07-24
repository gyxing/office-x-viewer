import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { PptParseError } from '../errors';

export type PptUserEdit = {
  offset: number;
  offsetLastEdit: number;
  offsetPersistDirectory: number;
  documentPersistId: number;
  persistIdSeed: number;
  encryptSessionPersistId?: number;
};

/** 从 PowerPoint Document 流的绝对偏移读取一次用户编辑记录。 */
export function readPptUserEdit(
  documentStream: Uint8Array,
  offset: number,
): PptUserEdit {
  const record = new PptRecordReader(
    documentStream,
    offset,
    documentStream.length,
  ).readRecord();
  if (!record || record.type !== PPT_RECORD.USER_EDIT_ATOM) {
    throw new PptParseError(
      'PPT_EDIT_CHAIN_INVALID',
      '编辑链指向的位置不是 UserEditAtom',
      { offset, recordType: record?.type },
    );
  }
  if (record.length !== 28 && record.length !== 32) {
    throw new PptParseError(
      'PPT_EDIT_CHAIN_INVALID',
      'UserEditAtom 长度无效',
      { offset, recordType: record.type },
    );
  }

  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  const offsetLastEdit = view.getUint32(8, true);
  const offsetPersistDirectory = view.getUint32(12, true);
  if (
    offsetLastEdit >= offset ||
    offsetPersistDirectory <= offsetLastEdit ||
    offsetPersistDirectory >= offset
  ) {
    throw new PptParseError(
      'PPT_EDIT_CHAIN_INVALID',
      'UserEditAtom 中的编辑或持久化目录偏移无效',
      { offset, recordType: record.type },
    );
  }

  return {
    offset,
    offsetLastEdit,
    offsetPersistDirectory,
    documentPersistId: view.getUint32(16, true),
    persistIdSeed: view.getUint32(20, true),
    encryptSessionPersistId:
      record.length === 32 ? view.getUint32(28, true) : undefined,
  };
}
