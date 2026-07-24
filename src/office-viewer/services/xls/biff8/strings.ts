import type { SpreadsheetWarning } from '../../spreadsheet/types';
import { XlsParseError } from '../errors';
import { Biff8Reader, type Biff8Record } from './Biff8Reader';

export type DecodedBiff8String = {
  value: string;
  bytesConsumed: number;
};

function decodeCharacters(bytes: Uint8Array, isWide: boolean) {
  let result = '';
  if (isWide) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    for (let offset = 0; offset < bytes.length; offset += 2) {
      result += String.fromCharCode(view.getUint16(offset, true));
    }
  } else {
    // BIFF8 压缩 Unicode 是低字节直映射，不是系统代码页文本。
    for (const value of bytes) result += String.fromCharCode(value);
  }
  return result;
}

/** 读取单个记录内的 BIFF8 XLUnicodeRichExtendedString。 */
export function readBiff8UnicodeString(
  reader: Biff8Reader,
  characterCountBytes: 1 | 2 = 2,
): DecodedBiff8String {
  const start = reader.position;
  const characterCount =
    characterCountBytes === 1 ? reader.readUint8() : reader.readUint16();
  const flags = reader.readUint8();
  const richRunCount = flags & 0x08 ? reader.readUint16() : 0;
  const extensionLength = flags & 0x04 ? reader.readUint32() : 0;
  const isWide = Boolean(flags & 0x01);
  const textBytes = reader.readBytes(characterCount * (isWide ? 2 : 1));
  const value = decodeCharacters(textBytes, isWide);
  reader.readBytes(richRunCount * 4);
  reader.readBytes(extensionLength);
  return { value, bytesConsumed: reader.position - start };
}

const CODE_PAGE_LABELS: Record<number, string> = {
  1252: 'windows-1252',
  936: 'gbk',
  950: 'big5',
  932: 'shift_jis',
  949: 'euc-kr',
};

/** 仅供 BIFF 旧式字节字符串字段使用，Unicode 字段不走代码页。 */
export function decodeLegacyByteString(
  bytes: Uint8Array,
  codePage: number | undefined,
  warnings: SpreadsheetWarning[],
) {
  const label = codePage ? CODE_PAGE_LABELS[codePage] : undefined;
  if (!label) {
    warnings.push({
      code: 'UNSUPPORTED_CODEPAGE',
      message: `未识别代码页 ${codePage ?? 'unknown'}，已按 windows-1252 解码`,
    });
  }
  try {
    return new TextDecoder(label ?? 'windows-1252').decode(bytes);
  } catch {
    return decodeCharacters(bytes, false);
  }
}

class SstSegmentReader {
  private segmentIndex = 0;
  private offset = 0;

  constructor(private readonly segments: Uint8Array[]) {}

  private moveToReadableSegment() {
    while (
      this.segmentIndex < this.segments.length &&
      this.offset >= this.segments[this.segmentIndex].length
    ) {
      this.segmentIndex += 1;
      this.offset = 0;
    }
    if (this.segmentIndex >= this.segments.length) {
      throw new Error('SST 字符串数据被截断');
    }
  }

  private readCurrentByte() {
    this.moveToReadableSegment();
    return this.segments[this.segmentIndex][this.offset++];
  }

  readUint8() {
    return this.readCurrentByte();
  }

  readUint16() {
    return this.readUint8() | (this.readUint8() << 8);
  }

  readUint32() {
    return (
      (this.readUint8() |
        (this.readUint8() << 8) |
        (this.readUint8() << 16) |
        (this.readUint8() << 24)) >>>
      0
    );
  }

  skip(length: number) {
    for (let index = 0; index < length; index += 1) this.readUint8();
  }

  readCharacters(characterCount: number, initialWide: boolean) {
    let value = '';
    let isWide = initialWide;
    for (let index = 0; index < characterCount; index += 1) {
      const segment = this.segments[this.segmentIndex];
      if (!segment || this.offset >= segment.length) {
        this.segmentIndex += 1;
        this.offset = 0;
        const continuation = this.readCurrentByte();
        isWide = Boolean(continuation & 0x01);
      }
      const current = this.segments[this.segmentIndex];
      const required = isWide ? 2 : 1;
      if (!current || this.offset + required > current.length) {
        throw new Error('SST 字符在 CONTINUE 边界处被截断');
      }
      const low = current[this.offset++];
      const high = isWide ? current[this.offset++] : 0;
      value += String.fromCharCode(low | (high << 8));
    }
    return value;
  }
}

/** 解析 SST 与连续 CONTINUE 记录，正确处理字符编码切换字节。 */
export function readBiff8SharedStrings(records: Biff8Record[]) {
  const first = records[0];
  if (!first || first.size < 8) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'BIFF8 SST 记录无效', {
      offset: first?.offset,
      recordId: first?.id,
    });
  }

  const header = new Biff8Reader(first.data);
  const totalCount = header.readUint32();
  const uniqueCount = header.readUint32();
  const segments = [
    first.data.subarray(8),
    ...records.slice(1).map((record) => record.data),
  ];
  const reader = new SstSegmentReader(segments);
  const strings: string[] = [];

  try {
    for (let index = 0; index < uniqueCount; index += 1) {
      const characterCount = reader.readUint16();
      const flags = reader.readUint8();
      const richRunCount = flags & 0x08 ? reader.readUint16() : 0;
      const extensionLength = flags & 0x04 ? reader.readUint32() : 0;
      strings.push(reader.readCharacters(characterCount, Boolean(flags & 1)));
      reader.skip(richRunCount * 4);
      reader.skip(extensionLength);
    }
  } catch (error) {
    const detail = error instanceof Error ? `：${error.message}` : '';
    throw new XlsParseError(
      'TRUNCATED_RECORD',
      `BIFF8 SST 字符串数据被截断${detail}`,
      { offset: first.offset, recordId: first.id },
    );
  }

  return { strings, totalCount, uniqueCount };
}
