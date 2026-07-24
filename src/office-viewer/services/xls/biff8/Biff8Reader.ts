import { XlsParseError } from '../errors';

export type Biff8Record = {
  id: number;
  offset: number;
  dataOffset: number;
  size: number;
  data: Uint8Array;
};

/** 对指定 BIFF 字节区间提供统一边界检查的读取器。 */
export class Biff8Reader {
  private readonly bytes: Uint8Array;
  private readonly start: number;
  private readonly end: number;
  private cursor: number;

  constructor(bytes: Uint8Array, start = 0, end = bytes.length) {
    if (start < 0 || end < start || end > bytes.length) {
      throw new XlsParseError('INVALID_RECORD_DATA', 'BIFF 读取区间无效');
    }
    this.bytes = bytes;
    this.start = start;
    this.end = end;
    this.cursor = start;
  }

  get position() {
    return this.cursor;
  }

  get remaining() {
    return this.end - this.cursor;
  }

  private ensureAvailable(length: number) {
    if (
      !Number.isInteger(length) ||
      length < 0 ||
      this.cursor + length > this.end
    ) {
      throw new XlsParseError('TRUNCATED_RECORD', 'BIFF 记录数据被截断', {
        offset: this.cursor,
      });
    }
  }

  readUint8() {
    this.ensureAvailable(1);
    return this.bytes[this.cursor++];
  }

  readUint16() {
    this.ensureAvailable(2);
    const value = this.bytes[this.cursor] | (this.bytes[this.cursor + 1] << 8);
    this.cursor += 2;
    return value;
  }

  readInt16() {
    const value = this.readUint16();
    return value & 0x8000 ? value - 0x10000 : value;
  }

  readUint32() {
    this.ensureAvailable(4);
    const view = new DataView(
      this.bytes.buffer,
      this.bytes.byteOffset + this.cursor,
      4,
    );
    this.cursor += 4;
    return view.getUint32(0, true);
  }

  readInt32() {
    this.ensureAvailable(4);
    const view = new DataView(
      this.bytes.buffer,
      this.bytes.byteOffset + this.cursor,
      4,
    );
    this.cursor += 4;
    return view.getInt32(0, true);
  }

  readFloat64() {
    this.ensureAvailable(8);
    const view = new DataView(
      this.bytes.buffer,
      this.bytes.byteOffset + this.cursor,
      8,
    );
    this.cursor += 8;
    return view.getFloat64(0, true);
  }

  readBytes(length: number) {
    this.ensureAvailable(length);
    const result = this.bytes.subarray(this.cursor, this.cursor + length);
    this.cursor += length;
    return result;
  }

  seek(position: number) {
    if (
      !Number.isInteger(position) ||
      position < this.start ||
      position > this.end
    ) {
      throw new XlsParseError('INVALID_RECORD_DATA', 'BIFF seek 位置无效', {
        offset: position,
      });
    }
    this.cursor = position;
  }
}

/** 顺序遍历 BIFF 记录，peek 不改变当前位置。 */
export class Biff8RecordCursor {
  private readonly stream: Uint8Array;
  private readonly end: number;
  private cursor: number;

  constructor(stream: Uint8Array, start = 0, end = stream.length) {
    if (start < 0 || end < start || end > stream.length) {
      throw new XlsParseError('INVALID_RECORD_DATA', 'BIFF 子流范围无效');
    }
    this.stream = stream;
    this.cursor = start;
    this.end = end;
  }

  get position() {
    return this.cursor;
  }

  peek(): Biff8Record | undefined {
    if (this.cursor === this.end) return undefined;
    if (this.cursor + 4 > this.end) {
      throw new XlsParseError('TRUNCATED_RECORD', 'BIFF 记录头被截断', {
        offset: this.cursor,
      });
    }
    const reader = new Biff8Reader(this.stream, this.cursor, this.end);
    const id = reader.readUint16();
    const size = reader.readUint16();
    const dataOffset = this.cursor + 4;
    if (dataOffset + size > this.end) {
      throw new XlsParseError('TRUNCATED_RECORD', 'BIFF 记录负载被截断', {
        offset: this.cursor,
        recordId: id,
      });
    }
    return {
      id,
      offset: this.cursor,
      dataOffset,
      size,
      data: this.stream.subarray(dataOffset, dataOffset + size),
    };
  }

  next() {
    const record = this.peek();
    if (record) this.cursor = record.dataOffset + record.size;
    return record;
  }
}

export type ParseYieldState = {
  lastYieldAt: number;
  budgetMs: number;
};

function currentTime() {
  return typeof performance !== 'undefined' ? performance.now() : Date.now();
}

/** 创建主线程解析的协作式时间片状态。 */
export function createParseYieldState(budgetMs = 8): ParseYieldState {
  return {
    lastYieldAt: currentTime(),
    budgetMs: Math.max(1, budgetMs),
  };
}

/** 时间片耗尽时让出一次浏览器事件循环，不改变记录处理顺序。 */
export async function yieldToBrowserIfNeeded(state: ParseYieldState) {
  const now = currentTime();
  if (now - state.lastYieldAt < state.budgetMs) return;
  await new Promise<void>((resolve) => {
    setTimeout(resolve, 0);
  });
  state.lastYieldAt = currentTime();
}
