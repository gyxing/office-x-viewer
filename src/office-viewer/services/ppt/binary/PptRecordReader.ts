import { PptParseError } from '../errors';
import type { PptRecord } from '../types';

function assertRange(
  offset: number,
  length: number,
  start: number,
  end: number,
) {
  if (
    !Number.isSafeInteger(offset) ||
    !Number.isSafeInteger(length) ||
    length < 0 ||
    offset < start ||
    offset > end ||
    length > end - offset
  ) {
    throw new PptParseError(
      'PPT_RECORD_OUT_OF_RANGE',
      'PowerPoint 记录范围超出数据流边界',
      { offset },
    );
  }
}

/** 在指定字节边界内顺序读取 PowerPoint 二进制记录。 */
export class PptRecordReader {
  private readonly view: DataView;
  private position: number;

  constructor(
    private readonly bytes: Uint8Array,
    private readonly start = 0,
    private readonly end = bytes.length,
  ) {
    assertRange(start, end - start, 0, bytes.length);
    this.position = start;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  get offset() {
    return this.position;
  }

  get remaining() {
    return this.end - this.position;
  }

  seek(offset: number) {
    assertRange(offset, 0, this.start, this.end);
    this.position = offset;
  }

  readRecord(): PptRecord | undefined {
    if (this.position === this.end) return undefined;
    if (this.remaining < 8) {
      throw new PptParseError(
        'PPT_RECORD_TRUNCATED',
        'PowerPoint 记录头不完整',
        { offset: this.position },
      );
    }

    const offset = this.position;
    const versionAndInstance = this.view.getUint16(offset, true);
    const type = this.view.getUint16(offset + 2, true);
    const length = this.view.getUint32(offset + 4, true);
    const dataOffset = offset + 8;
    assertRange(dataOffset, length, this.start, this.end);
    const endOffset = dataOffset + length;
    this.position = endOffset;

    return {
      version: versionAndInstance & 0x000f,
      instance: versionAndInstance >>> 4,
      type,
      length,
      offset,
      dataOffset,
      endOffset,
      data: this.bytes.subarray(dataOffset, endOffset),
    };
  }

  peekRecord() {
    const offset = this.position;
    const record = this.readRecord();
    this.position = offset;
    return record;
  }

  childReader(record: PptRecord) {
    return new PptRecordReader(this.bytes, record.dataOffset, record.endOffset);
  }

  *records(): IterableIterator<PptRecord> {
    let record = this.readRecord();
    while (record) {
      yield record;
      record = this.readRecord();
    }
  }
}

/** 主线程解析超过一个时间片后让出浏览器事件循环。 */
export function createPptTimeSlice(budgetMs = 8) {
  let startedAt = Date.now();
  return async () => {
    if (Date.now() - startedAt < budgetMs) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    startedAt = Date.now();
  };
}
