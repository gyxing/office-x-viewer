import { OFFICE_ART_RECORD } from './constants';
import { OfficeArtParseError } from './OfficeArtParseError';
import type { OfficeArtRecord, OfficeArtWarning } from './types';

const KNOWN_TYPES = new Set<number>(Object.values(OFFICE_ART_RECORD));

function parseRecordRange(
  bytes: Uint8Array,
  start: number,
  end: number,
  warnings: OfficeArtWarning[],
  unknownTypes: Set<number>,
) {
  const records: OfficeArtRecord[] = [];
  let offset = start;
  while (offset < end) {
    if (offset + 8 > end) {
      throw new OfficeArtParseError(
        'OfficeArt 记录头超出父容器边界',
        offset,
      );
    }
    const view = new DataView(
      bytes.buffer,
      bytes.byteOffset + offset,
      end - offset,
    );
    const options = view.getUint16(0, true);
    const version = options & 0x000f;
    const instance = options >> 4;
    const type = view.getUint16(2, true);
    const length = view.getUint32(4, true);
    const dataOffset = offset + 8;
    const recordEnd = dataOffset + length;
    if (
      !Number.isSafeInteger(recordEnd) ||
      recordEnd < dataOffset ||
      recordEnd > end
    ) {
      throw new OfficeArtParseError(
        `OfficeArt 记录 0x${type.toString(16)} 长度越界`,
        offset,
      );
    }
    if (
      type >= 0xf000 &&
      !KNOWN_TYPES.has(type) &&
      !unknownTypes.has(type)
    ) {
      unknownTypes.add(type);
      warnings.push({
        code: 'UNKNOWN_OFFICE_ART_RECORD',
        message: `已跳过未知 OfficeArt 记录 0x${type
          .toString(16)
          .toUpperCase()}`,
        offset,
      });
    }
    const data = bytes.subarray(dataOffset, recordEnd);
    records.push({
      version,
      instance,
      type,
      length,
      offset,
      data,
      children:
        version === 0x0f
          ? parseRecordRange(
              bytes,
              dataOffset,
              recordEnd,
              warnings,
              unknownTypes,
            )
          : undefined,
    });
    offset = recordEnd;
  }
  return records;
}

/** 严格按父容器边界解析可由 XLS 与 PPT 复用的 OfficeArt 记录树。 */
export function parseOfficeArtRecords(
  bytes: Uint8Array,
  warnings: OfficeArtWarning[] = [],
) {
  return parseRecordRange(bytes, 0, bytes.length, warnings, new Set<number>());
}
