import type { TextStyle } from '../../presentation/types';
import type { PptParseContext, PptRecord } from '../types';
import type {
  PptCharacterStyleRun,
  PptParagraphStyleRun,
  PptTextDefaults,
  PptTextStyleRuns,
} from './types';

const alignments: Array<TextStyle['align']> = [
  'left',
  'center',
  'right',
  'justify',
  'justify',
  'justify',
  'justify',
];

class StyleReader {
  private readonly view: DataView;

  offset = 0;

  constructor(private readonly bytes: Uint8Array) {
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  get remaining() {
    return this.bytes.length - this.offset;
  }

  u16() {
    if (this.remaining < 2) throw new RangeError('文本样式记录已截断');
    const value = this.view.getUint16(this.offset, true);
    this.offset += 2;
    return value;
  }

  i16() {
    if (this.remaining < 2) throw new RangeError('文本样式记录已截断');
    const value = this.view.getInt16(this.offset, true);
    this.offset += 2;
    return value;
  }

  u32() {
    if (this.remaining < 4) throw new RangeError('文本样式记录已截断');
    const value = this.view.getUint32(this.offset, true);
    this.offset += 4;
    return value;
  }

  skip(length: number) {
    if (this.remaining < length) throw new RangeError('文本样式记录已截断');
    this.offset += length;
  }
}

function readColor(value: number) {
  const red = value & 0xff;
  const green = (value >>> 8) & 0xff;
  const blue = (value >>> 16) & 0xff;
  return `#${[red, green, blue]
    .map((part) => part.toString(16).padStart(2, '0'))
    .join('')}`;
}

function readLineHeight(value: number) {
  // 正值是行高百分比；负值是绝对主单位，统一模型暂无绝对行高字段。
  return value >= 0 ? Math.max(0.1, value / 100) : undefined;
}

function readParagraphSpacing(value: number) {
  // 正值按默认 18px 行高折算，负值按 8 主单位 = 1px 折算。
  return value >= 0 ? (value / 100) * 18 : Math.abs(value) / 8;
}

function readParagraphException(reader: StyleReader) {
  const masks = reader.u32();
  const style: TextStyle = {};
  let bulletFlags: number | undefined;
  if (masks & 0x0f) bulletFlags = reader.u16();
  let bulletChar: number | undefined;
  if (masks & (1 << 7)) bulletChar = reader.u16();
  if (masks & (1 << 4)) reader.u16();
  if (masks & (1 << 6)) {
    const size = reader.i16();
    if (size > 0) style.bullet = { ...style.bullet, size };
  }
  if (masks & (1 << 5)) {
    style.bullet = { ...style.bullet, color: readColor(reader.u32()) };
  }
  if (masks & (1 << 11)) {
    style.align = alignments[reader.u16()] ?? 'left';
  }
  if (masks & (1 << 12)) style.lineHeight = readLineHeight(reader.i16());
  if (masks & (1 << 13)) {
    style.spaceBefore = readParagraphSpacing(reader.i16());
  }
  if (masks & (1 << 14)) {
    style.spaceAfter = readParagraphSpacing(reader.i16());
  }
  if (masks & (1 << 8)) style.marginLeft = reader.i16() / 8;
  if (masks & (1 << 10)) style.textIndent = reader.i16() / 8;
  if (masks & (1 << 15)) reader.u16();
  if (masks & (1 << 20)) {
    const tabCount = reader.u16();
    reader.skip(tabCount * 4);
  }
  if (masks & (1 << 16)) reader.u16();
  if (masks & ((1 << 17) | (1 << 18) | (1 << 19))) reader.u16();
  if (masks & (1 << 21)) {
    style.writingMode = reader.u16() === 1 ? 'vertical-rl' : 'horizontal-tb';
  }
  if (bulletFlags !== undefined) {
    const hasBullet = Boolean(bulletFlags & 1);
    style.bullet = hasBullet
      ? { ...style.bullet, char: bulletChar ? String.fromCharCode(bulletChar) : '•' }
      : { none: true };
  }
  return style;
}

function readCharacterException(
  reader: StyleReader,
  defaults: PptTextDefaults,
) {
  const masks = reader.u32();
  const style: TextStyle = {};
  if (masks & 0xffff) {
    const flags = reader.u16();
    if (masks & 1) style.bold = Boolean(flags & 1);
    if (masks & 2) style.italic = Boolean(flags & 2);
    if (masks & 4) style.underline = Boolean(flags & 4);
  }
  if (masks & (1 << 16)) {
    const fontRef = reader.u16();
    style.fontFamily = defaults.fonts?.get(fontRef);
  }
  if (masks & (1 << 21)) reader.u16();
  if (masks & (1 << 22)) {
    const ansiFontRef = reader.u16();
    style.fontFamily = defaults.fonts?.get(ansiFontRef) ?? style.fontFamily;
  }
  if (masks & (1 << 23)) reader.u16();
  if (masks & (1 << 17)) style.fontSize = (reader.u16() * 4) / 3;
  if (masks & (1 << 18)) style.color = readColor(reader.u32());
  if (masks & (1 << 19)) style.baseline = reader.i16();
  return style;
}

function addWarning(
  context: PptParseContext,
  record: PptRecord,
  message: string,
) {
  context.warnings.push({
    code: 'PPT_TEXT_RUN_TRUNCATED',
    message,
    offset: record.offset,
  });
}

/** 解码 StyleTextPropAtom 的段落运行与字符运行。 */
export function readPptTextStyles(
  record: PptRecord | undefined,
  textLength: number,
  defaults: PptTextDefaults,
  context: PptParseContext,
): PptTextStyleRuns {
  const result: PptTextStyleRuns = { paragraphs: [], characters: [] };
  if (!record) return result;
  const reader = new StyleReader(record.data);
  const targetLength = textLength + 1;

  try {
    let covered = 0;
    while (covered < targetLength && reader.remaining >= 10) {
      const count = reader.u32();
      const level = reader.u16();
      const style = readParagraphException(reader);
      result.paragraphs.push({ count, level, style });
      covered += count;
      if (!count) break;
    }
    if (covered < textLength) {
      addWarning(context, record, '段落样式运行长度小于文本长度');
    }

    covered = 0;
    while (covered < targetLength && reader.remaining >= 8) {
      const count = reader.u32();
      const style = readCharacterException(reader, defaults);
      result.characters.push({ count, style });
      covered += count;
      if (!count) break;
    }
    if (covered < textLength) {
      addWarning(context, record, '字符样式运行长度小于文本长度');
    }
  } catch (error) {
    addWarning(
      context,
      record,
      error instanceof Error ? error.message : '文本样式记录无法完整读取',
    );
  }
  return result;
}

export type { PptCharacterStyleRun, PptParagraphStyleRun };
