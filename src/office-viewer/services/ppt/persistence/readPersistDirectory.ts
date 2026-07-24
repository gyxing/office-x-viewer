import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { PptParseError } from '../errors';

/** 读取一次编辑对应的压缩 persist ID 到流偏移映射。 */
export function readPptPersistDirectory(
  documentStream: Uint8Array,
  offset: number,
  lastEditOffset: number,
) {
  const record = new PptRecordReader(
    documentStream,
    offset,
    documentStream.length,
  ).readRecord();
  if (
    !record ||
    (record.type !== PPT_RECORD.PERSIST_PTR_FULL_BLOCK &&
      record.type !== PPT_RECORD.PERSIST_PTR_INCREMENTAL_BLOCK)
  ) {
    throw new PptParseError(
      'PPT_PERSIST_DIRECTORY_INVALID',
      '编辑链指向的位置不是 PersistDirectoryAtom',
      { offset, recordType: record?.type },
    );
  }

  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  const result = new Map<number, number>();
  let cursor = 0;
  while (cursor < record.data.length) {
    if (record.data.length - cursor < 4) {
      throw new PptParseError(
        'PPT_PERSIST_DIRECTORY_INVALID',
        'PersistDirectoryEntry 头不完整',
        { offset: record.dataOffset + cursor, recordType: record.type },
      );
    }
    const descriptor = view.getUint32(cursor, true);
    cursor += 4;
    const persistId = descriptor & 0x000fffff;
    const count = descriptor >>> 20;
    if (!count || count > (record.data.length - cursor) / 4) {
      throw new PptParseError(
        'PPT_PERSIST_DIRECTORY_INVALID',
        'PersistDirectoryEntry 数量超出记录边界',
        { offset: record.dataOffset + cursor - 4, recordType: record.type },
      );
    }

    for (let index = 0; index < count; index += 1) {
      const persistOffset = view.getUint32(cursor, true);
      cursor += 4;
      if (
        persistOffset < lastEditOffset ||
        persistOffset >= offset ||
        persistOffset >= documentStream.length
      ) {
        throw new PptParseError(
          'PPT_PERSIST_DIRECTORY_INVALID',
          'Persist 对象偏移超出当前编辑范围',
          { offset: record.dataOffset + cursor - 4, recordType: record.type },
        );
      }
      result.set(persistId + index, persistOffset);
    }
  }
  return result;
}
