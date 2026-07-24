import { PPT_RECORD } from '../binary/constants';
import {
  readPptByteString,
  readPptUnicodeString,
} from '../binary/readStrings';
import type { PptParseContext, PptRecord } from '../types';
import type { PptTextAtomGroup } from './types';

function decodeText(record: PptRecord, context: PptParseContext) {
  const value =
    record.type === PPT_RECORD.TEXT_CHARS_ATOM
      ? readPptUnicodeString(record.data)
      : readPptByteString(record.data, 1252, context);
  return value.replace(/\u0000+$/g, '');
}

/** 将 TextHeader 与其内容、样式记录关联为可独立恢复的文本组。 */
export function readPptTextAtoms(
  records: PptRecord[],
  context: PptParseContext,
): PptTextAtomGroup[] {
  const groups: PptTextAtomGroup[] = [];
  let textType = 4;

  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];
    if (record.type === PPT_RECORD.TEXT_HEADER_ATOM) {
      if (record.length >= 4) {
        textType = new DataView(
          record.data.buffer,
          record.data.byteOffset,
          record.data.byteLength,
        ).getUint32(0, true);
      }
      continue;
    }
    if (
      record.type !== PPT_RECORD.TEXT_CHARS_ATOM &&
      record.type !== PPT_RECORD.TEXT_BYTES_ATOM
    ) {
      continue;
    }

    const styleRecord =
      records[index + 1]?.type === PPT_RECORD.STYLE_TEXT_PROP_ATOM
        ? records[index + 1]
        : undefined;
    groups.push({
      textType,
      text: decodeText(record, context),
      contentRecord: record,
      styleRecord,
    });
  }
  return groups;
}
