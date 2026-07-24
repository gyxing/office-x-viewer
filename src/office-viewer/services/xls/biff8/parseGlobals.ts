import type { SpreadsheetWarning } from '../../spreadsheet/types';
import { XlsParseError } from '../errors';
import type {
  Biff8CellFormat,
  Biff8Font,
  Biff8SheetDescriptor,
  Biff8SheetType,
  Biff8WorkbookGlobals,
} from '../types';
import {
  Biff8Reader,
  Biff8RecordCursor,
  yieldToBrowserIfNeeded,
  type Biff8Record,
  type ParseYieldState,
} from './Biff8Reader';
import {
  BIFF8_RECORD,
  BIFF8_SUBSTREAM,
  BIFF8_VERSION,
  DEFAULT_BIFF8_PALETTE,
} from './constants';
import { BUILTIN_NUMBER_FORMATS } from './numberFormats';
import { readBiff8SharedStrings, readBiff8UnicodeString } from './strings';

function validateGlobalsBof(stream: Uint8Array) {
  const cursor = new Biff8RecordCursor(stream);
  const record = cursor.next();
  if (!record || record.id !== BIFF8_RECORD.BOF || record.size < 4) {
    throw new XlsParseError(
      'INVALID_RECORD_DATA',
      'Workbook Globals 缺少有效 BOF',
      { offset: record?.offset, recordId: record?.id },
    );
  }
  const reader = new Biff8Reader(record.data);
  const version = reader.readUint16();
  const substreamType = reader.readUint16();
  if (version !== BIFF8_VERSION) {
    throw new XlsParseError(
      'UNSUPPORTED_BIFF_VERSION',
      `仅支持 BIFF8，当前版本为 0x${version.toString(16)}`,
      { offset: record.offset, recordId: record.id },
    );
  }
  if (substreamType !== BIFF8_SUBSTREAM.WORKBOOK_GLOBALS) {
    throw new XlsParseError(
      'INVALID_RECORD_DATA',
      'Workbook 流首个 BOF 不是 Globals 子流',
      { offset: record.offset, recordId: record.id },
    );
  }
  return cursor;
}

function sheetTypeFromValue(value: number): Biff8SheetType {
  if (value === 0x00) return 'worksheet';
  if (value === 0x02) return 'chart';
  if (value === 0x01) return 'macro';
  if (value === 0x06) return 'dialog';
  return 'unknown';
}

function parseBoundSheet(record: Biff8Record): Biff8SheetDescriptor {
  const reader = new Biff8Reader(record.data);
  const streamOffset = reader.readUint32();
  const visibilityValue = reader.readUint8();
  const type = sheetTypeFromValue(reader.readUint8());
  const name = readBiff8UnicodeString(reader, 1).value;
  return {
    id: `xls-sheet-${record.offset}`,
    name,
    streamOffset,
    visibility:
      visibilityValue === 1
        ? 'hidden'
        : visibilityValue === 2
        ? 'veryHidden'
        : 'visible',
    type,
  };
}

function parseFont(data: Uint8Array): Biff8Font {
  const reader = new Biff8Reader(data);
  const heightTwips = reader.readUint16();
  const flags = reader.readUint16();
  const colorIndex = reader.readUint16();
  const boldWeight = reader.readUint16();
  reader.readBytes(6);
  const name = readBiff8UnicodeString(reader, 1).value;
  return {
    name,
    heightTwips,
    colorIndex,
    bold: boldWeight >= 700,
    italic: Boolean(flags & 0x0002),
    underline: data[10] !== 0,
  };
}

function parseCellFormat(data: Uint8Array): Biff8CellFormat {
  const reader = new Biff8Reader(data);
  const fontIndex = reader.readUint16();
  const formatIndex = reader.readUint16();
  const protection = reader.readUint16();
  const alignment = reader.readUint8();
  reader.readBytes(3);
  const borderStyles = reader.readUint16();
  const borderColors = reader.readUint16();
  const borderAndFill = reader.readUint32();
  const fillColors = reader.readUint16();
  return {
    fontIndex,
    formatIndex,
    parentStyleIndex: (protection >> 4) & 0x0fff,
    isStyle: Boolean(protection & 0x0004),
    horizontalAlign: alignment & 0x07,
    verticalAlign: (alignment >> 4) & 0x07,
    wrapText: Boolean(alignment & 0x08),
    fillPattern: (borderAndFill >>> 26) & 0x3f,
    fillForegroundColorIndex: fillColors & 0x7f,
    fillBackgroundColorIndex: (fillColors >> 7) & 0x7f,
    leftBorder: {
      style: borderStyles & 0x0f,
      colorIndex: borderColors & 0x7f,
    },
    rightBorder: {
      style: (borderStyles >> 4) & 0x0f,
      colorIndex: (borderColors >> 7) & 0x7f,
    },
    topBorder: {
      style: (borderStyles >> 8) & 0x0f,
      colorIndex: borderAndFill & 0x7f,
    },
    bottomBorder: {
      style: (borderStyles >> 12) & 0x0f,
      colorIndex: (borderAndFill >> 7) & 0x7f,
    },
  };
}

function parsePalette(data: Uint8Array) {
  const reader = new Biff8Reader(data);
  const count = reader.readUint16();
  const palette: string[] = [];
  for (let index = 0; index < count; index += 1) {
    const red = reader.readUint8();
    const green = reader.readUint8();
    const blue = reader.readUint8();
    reader.readUint8();
    palette.push(
      `#${[red, green, blue]
        .map((value) => value.toString(16).padStart(2, '0'))
        .join('')}`,
    );
  }
  return palette;
}

function parseDefinedName(data: Uint8Array, id: number) {
  const header = new Biff8Reader(data);
  header.readUint16();
  header.readUint8();
  const characterCount = header.readUint8();
  const tokenLength = header.readUint16();
  header.readBytes(8);
  const isWide = Boolean(header.readUint8() & 0x01);
  const nameBytes = header.readBytes(characterCount * (isWide ? 2 : 1));
  let name = '';
  if (isWide) {
    const view = new DataView(
      nameBytes.buffer,
      nameBytes.byteOffset,
      nameBytes.byteLength,
    );
    for (let offset = 0; offset < nameBytes.length; offset += 2) {
      name += String.fromCharCode(view.getUint16(offset, true));
    }
  } else {
    for (const value of nameBytes) name += String.fromCharCode(value);
  }
  return {
    id,
    name,
    tokens: header.readBytes(tokenLength),
  };
}

function validateSheetOffsets(
  workbookStream: Uint8Array,
  sheets: Biff8SheetDescriptor[],
) {
  const offsets = new Set<number>();
  for (const sheet of sheets) {
    if (
      sheet.streamOffset < 0 ||
      sheet.streamOffset + 8 > workbookStream.length ||
      offsets.has(sheet.streamOffset)
    ) {
      throw new XlsParseError(
        'INVALID_RECORD_DATA',
        `工作表 ${sheet.name} 的子流偏移无效`,
        { offset: sheet.streamOffset },
      );
    }
    offsets.add(sheet.streamOffset);
    const record = new Biff8RecordCursor(
      workbookStream,
      sheet.streamOffset,
    ).next();
    if (!record || record.id !== BIFF8_RECORD.BOF || record.size < 4) {
      throw new XlsParseError(
        'INVALID_RECORD_DATA',
        `工作表 ${sheet.name} 的 BOF 无效`,
        { offset: sheet.streamOffset, recordId: record?.id },
      );
    }
    const version = new Biff8Reader(record.data).readUint16();
    if (version !== BIFF8_VERSION) {
      throw new XlsParseError(
        'UNSUPPORTED_BIFF_VERSION',
        `工作表 ${sheet.name} 不是 BIFF8 子流`,
        { offset: sheet.streamOffset, recordId: record.id },
      );
    }
  }
}

/** 解析 Workbook Globals、共享字符串、样式和工作表目录。 */
export async function parseBiff8Globals(
  workbookStream: Uint8Array,
  yieldState: ParseYieldState,
): Promise<Biff8WorkbookGlobals> {
  const cursor = validateGlobalsBof(workbookStream);
  const sheets: Biff8SheetDescriptor[] = [];
  const sharedStrings: string[] = [];
  const fonts: Biff8Font[] = [];
  const formats = new Map<number, string>(
    Object.entries(BUILTIN_NUMBER_FORMATS).map(([id, format]) => [
      Number(id),
      format,
    ]),
  );
  const cellFormats: Biff8CellFormat[] = [];
  const definedNames: Biff8WorkbookGlobals['definedNames'] = [];
  const warnings: SpreadsheetWarning[] = [];
  const drawingGroupRecords: Biff8WorkbookGlobals['drawingGroupRecords'] = [];
  let palette: string[] = [...DEFAULT_BIFF8_PALETTE];
  let date1904 = false;
  let codePage: number | undefined;
  let reachedEof = false;

  for (let record = cursor.next(); record; record = cursor.next()) {
    if (record.id === BIFF8_RECORD.EOF) {
      reachedEof = true;
      break;
    }
    if (record.id === BIFF8_RECORD.FILEPASS) {
      throw new XlsParseError(
        'ENCRYPTED_FILE',
        '暂不支持加密的 Excel 97–2003 文件',
        { offset: record.offset, recordId: record.id },
      );
    }

    switch (record.id) {
      case BIFF8_RECORD.CODEPAGE:
        codePage = new Biff8Reader(record.data).readUint16();
        break;
      case BIFF8_RECORD.BOUNDSHEET8:
        sheets.push(parseBoundSheet(record));
        break;
      case BIFF8_RECORD.FONT:
        fonts.push(parseFont(record.data));
        break;
      case BIFF8_RECORD.XF:
        cellFormats.push(parseCellFormat(record.data));
        break;
      case BIFF8_RECORD.PALETTE:
        palette = parsePalette(record.data);
        break;
      case BIFF8_RECORD.DATEMODE:
        date1904 = new Biff8Reader(record.data).readUint16() === 1;
        break;
      case BIFF8_RECORD.FORMAT: {
        const reader = new Biff8Reader(record.data);
        const formatId = reader.readUint16();
        formats.set(formatId, readBiff8UnicodeString(reader).value);
        break;
      }
      case BIFF8_RECORD.NAME:
        definedNames.push(
          parseDefinedName(record.data, definedNames.length + 1),
        );
        break;
      case BIFF8_RECORD.SST: {
        const records = [record];
        while (cursor.peek()?.id === BIFF8_RECORD.CONTINUE) {
          records.push(cursor.next()!);
        }
        const parsed = readBiff8SharedStrings(records);
        if (parsed.uniqueCount > parsed.totalCount) {
          throw new XlsParseError(
            'INVALID_RECORD_DATA',
            'SST 唯一字符串数量大于单元格引用总数',
            { offset: record.offset, recordId: record.id },
          );
        }
        sharedStrings.push(...parsed.strings);
        break;
      }
      case BIFF8_RECORD.MSODRAWINGGROUP: {
        const chunks = [record.data];
        while (cursor.peek()?.id === BIFF8_RECORD.CONTINUE) {
          chunks.push(cursor.next()!.data);
        }
        drawingGroupRecords.push({
          recordId: record.id,
          offset: record.offset,
          chunks,
        });
        break;
      }
      default:
        break;
    }
    await yieldToBrowserIfNeeded(yieldState);
  }

  if (!reachedEof) {
    throw new XlsParseError('TRUNCATED_RECORD', 'Workbook Globals 缺少 EOF');
  }
  validateSheetOffsets(workbookStream, sheets);
  for (const sheet of sheets) {
    if (sheet.type === 'macro' || sheet.type === 'dialog') {
      warnings.push({
        code: 'UNSUPPORTED_SHEET_TYPE',
        message: `工作表“${sheet.name}”属于宏表或对话表，当前不渲染`,
        sheetName: sheet.name,
        offset: sheet.streamOffset,
      });
    } else if (sheet.type === 'unknown') {
      warnings.push({
        code: 'UNSUPPORTED_SHEET_TYPE',
        message: `工作表“${sheet.name}”的类型暂不支持`,
        sheetName: sheet.name,
        offset: sheet.streamOffset,
      });
    }
  }

  return {
    sheets,
    sharedStrings,
    fonts,
    formats,
    cellFormats,
    palette,
    date1904,
    definedNames,
    warnings,
    hasVba: false,
    codePage,
    drawingGroupRecords,
  };
}
