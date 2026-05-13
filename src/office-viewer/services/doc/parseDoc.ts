import type {
  DocBlock,
  DocDocument,
  DocImage,
  DocListBlock,
  DocParagraph,
  DocParagraphBlock,
  DocTableBlock,
  DocTableStyle,
  DocTextInline,
  DocTextStyle,
} from './types';

type CfbDirectoryEntry = {
  name: string;
  objectType: number;
  startSector: number;
  streamSize: number;
};

type CfbFile = {
  streams: Map<string, Uint8Array>;
};

type DocPiece = {
  charStart: number;
  charEnd: number;
  fileOffset: number;
  compressed: boolean;
};

type DocFib = ReturnType<typeof parseFib>;

type DocCharacterRun = {
  fcStart: number;
  fcEnd: number;
  style: DocTextStyle;
};

type DocParagraphRun = {
  fcStart: number;
  fcEnd: number;
  style: DocTextStyle;
};

type DocTableRun = {
  fcStart: number;
  fcEnd: number;
  style: DocTableStyle;
};

type DocTextSegment = {
  text: string;
  style?: DocTextStyle;
};

type DocImageSegment = {
  text: string;
  style?: DocTextStyle;
  image?: DocImage;
};

type DocImageCandidate = DocImage & {
  offset: number;
  byteLength: number;
  packagedMedia: boolean;
  webExtensionPreview: boolean;
  streamName: string;
};

type ParsedListLine = {
  ordered: boolean;
  text: string;
  inlines?: DocTextInline[];
};

type DocLine = {
  text: string;
  inlines: DocTextInline[];
  style?: DocTextStyle;
  match: (regexp: RegExp) => RegExpMatchArray | null;
};

type PendingTableCell = {
  text: string;
  inlines: DocTextInline[];
  style?: DocTextStyle;
};

type DocFontTable = string[];

// 旧版 .doc 是 OLE/CFB 二进制容器，不是 zip；这里实现最小可用的前端降级解析。
const DOC_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREE_SECTOR = 0xffffffff;
const END_OF_CHAIN = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const MINI_STREAM_CUTOFF_SIZE = 4096;

const DEFAULT_DOC_PAGE = {
  width: 794,
  minHeight: 1123,
  marginTop: 96,
  marginRight: 120,
  marginBottom: 96,
  marginLeft: 120,
};

const DOC_FONT_FAMILY = '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif';
const WORD_ICO_COLORS: Record<number, string> = {
  1: '#000000',
  2: '#0000ff',
  3: '#00ffff',
  4: '#00ff00',
  5: '#ff00ff',
  6: '#ff0000',
  7: '#ffff00',
  8: '#ffffff',
  9: '#000080',
  10: '#008080',
  11: '#008000',
  12: '#800080',
  13: '#800000',
  14: '#808000',
  15: '#808080',
  16: '#c0c0c0',
};

function isBlobInput(file: File | Blob | ArrayBuffer | Uint8Array): file is File | Blob {
  return typeof Blob !== 'undefined' && file instanceof Blob;
}

async function readBytes(file: File | Blob | ArrayBuffer | Uint8Array) {
  if (file instanceof Uint8Array) return file;
  const buffer = isBlobInput(file) ? await file.arrayBuffer() : file;
  return new Uint8Array(buffer);
}

function isOleDoc(bytes: Uint8Array) {
  return DOC_MAGIC.every((value, index) => bytes[index] === value);
}

function readUint16(view: DataView, offset: number) {
  return view.getUint16(offset, true);
}

function readUint32(view: DataView, offset: number) {
  return view.getUint32(offset, true);
}

function readUint16BE(view: DataView, offset: number) {
  return view.getUint16(offset, false);
}

function readUint32BE(view: DataView, offset: number) {
  return view.getUint32(offset, false);
}

function readInt16(view: DataView, offset: number) {
  return view.getInt16(offset, true);
}

function twipToPx(value: number) {
  return (value / 1440) * 96;
}

function sectorOffset(sector: number, sectorSize: number) {
  return (sector + 1) * sectorSize;
}

function sliceSector(bytes: Uint8Array, sector: number, sectorSize: number) {
  const offset = sectorOffset(sector, sectorSize);
  return bytes.slice(offset, offset + sectorSize);
}

function concatChunks(chunks: Uint8Array[]) {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;

  chunks.forEach((chunk) => {
    result.set(chunk, offset);
    offset += chunk.length;
  });

  return result;
}

function readSectorChain(startSector: number, fat: number[], bytes: Uint8Array, sectorSize: number) {
  const chunks: Uint8Array[] = [];
  const seen = new Set<number>();
  let sector = startSector;

  while (sector !== END_OF_CHAIN && sector !== FREE_SECTOR && sector < fat.length && !seen.has(sector)) {
    seen.add(sector);
    chunks.push(sliceSector(bytes, sector, sectorSize));
    sector = fat[sector] ?? END_OF_CHAIN;
  }

  return concatChunks(chunks);
}

function readMiniSectorChain(startSector: number, miniFat: number[], miniStream: Uint8Array, miniSectorSize: number) {
  const chunks: Uint8Array[] = [];
  const seen = new Set<number>();
  let sector = startSector;

  while (sector !== END_OF_CHAIN && sector !== FREE_SECTOR && sector < miniFat.length && !seen.has(sector)) {
    seen.add(sector);
    const offset = sector * miniSectorSize;
    chunks.push(miniStream.slice(offset, offset + miniSectorSize));
    sector = miniFat[sector] ?? END_OF_CHAIN;
  }

  return concatChunks(chunks);
}

function decodeUtf16Name(bytes: Uint8Array, length: number) {
  const chars: number[] = [];
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const max = Math.max(0, length - 2);

  for (let offset = 0; offset < max; offset += 2) {
    chars.push(view.getUint16(offset, true));
  }

  return String.fromCharCode(...chars);
}

function parseDirectoryEntries(directoryStream: Uint8Array) {
  const entries: CfbDirectoryEntry[] = [];

  for (let offset = 0; offset + 128 <= directoryStream.length; offset += 128) {
    const entryBytes = directoryStream.slice(offset, offset + 128);
    const view = new DataView(entryBytes.buffer, entryBytes.byteOffset, entryBytes.byteLength);
    const nameLength = readUint16(view, 64);
    const name = decodeUtf16Name(entryBytes.slice(0, 64), nameLength);
    const objectType = entryBytes[66];
    const startSector = readUint32(view, 116);
    const streamSize = readUint32(view, 120);

    if (name && objectType !== 0) {
      entries.push({ name, objectType, startSector, streamSize });
    }
  }

  return entries;
}

function readDifatSectorEntries(bytes: Uint8Array, sector: number, sectorSize: number) {
  const data = sliceSector(bytes, sector, sectorSize);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const values: number[] = [];
  const entryCount = Math.floor((sectorSize - 4) / 4);

  for (let index = 0; index < entryCount; index += 1) {
    values.push(readUint32(view, index * 4));
  }

  return {
    values,
    nextSector: readUint32(view, sectorSize - 4),
  };
}

function readFat(bytes: Uint8Array, sectorSize: number, difat: number[]) {
  const fat: number[] = [];

  difat.forEach((sector) => {
    if (sector === FREE_SECTOR || sector === END_OF_CHAIN || sector === FAT_SECTOR) return;
    const data = sliceSector(bytes, sector, sectorSize);
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    for (let offset = 0; offset + 4 <= data.length; offset += 4) {
      fat.push(readUint32(view, offset));
    }
  });

  return fat;
}

function readMiniFat(bytes: Uint8Array, startSector: number, sectorCount: number, fat: number[], sectorSize: number) {
  if (!sectorCount || startSector === END_OF_CHAIN) return [];

  const data = readSectorChain(startSector, fat, bytes, sectorSize).slice(0, sectorCount * sectorSize);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const miniFat: number[] = [];

  for (let offset = 0; offset + 4 <= data.length; offset += 4) {
    miniFat.push(readUint32(view, offset));
  }

  return miniFat;
}

function parseCfb(bytes: Uint8Array): CfbFile {
  // CFB 先通过 FAT/miniFAT 还原各个 stream，后续 WordDocument/Table stream 才能继续解析。
  if (!isOleDoc(bytes)) {
    throw new Error('\u4e0d\u662f\u6709\u6548\u7684 Word 97-2003 DOC \u6587\u4ef6');
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorSize = 2 ** readUint16(view, 30);
  const miniSectorSize = 2 ** readUint16(view, 32);
  const directoryStartSector = readUint32(view, 48);
  const miniFatStartSector = readUint32(view, 60);
  const miniFatSectorCount = readUint32(view, 64);
  const difatStartSector = readUint32(view, 68);
  const difatSectorCount = readUint32(view, 72);
  const difat: number[] = [];

  for (let offset = 76; offset < 512; offset += 4) {
    const value = readUint32(view, offset);
    if (value !== FREE_SECTOR) difat.push(value);
  }

  let nextDifatSector = difatStartSector;
  for (let index = 0; index < difatSectorCount && nextDifatSector !== END_OF_CHAIN; index += 1) {
    const sector = readDifatSectorEntries(bytes, nextDifatSector, sectorSize);
    sector.values.forEach((value) => {
      if (value !== FREE_SECTOR) difat.push(value);
    });
    nextDifatSector = sector.nextSector;
  }

  const fat = readFat(bytes, sectorSize, difat);
  const directoryStream = readSectorChain(directoryStartSector, fat, bytes, sectorSize);
  const entries = parseDirectoryEntries(directoryStream);
  const root = entries.find((entry) => entry.objectType === 5);
  const miniStream =
    root && root.startSector !== END_OF_CHAIN
      ? readSectorChain(root.startSector, fat, bytes, sectorSize).slice(0, root.streamSize)
      : new Uint8Array();
  const miniFat = readMiniFat(bytes, miniFatStartSector, miniFatSectorCount, fat, sectorSize);
  const streams = new Map<string, Uint8Array>();

  entries
    .filter((entry) => entry.objectType === 2)
    .forEach((entry) => {
      const isMiniStream = entry.streamSize < MINI_STREAM_CUTOFF_SIZE && entry.startSector !== END_OF_CHAIN;
      const data = isMiniStream
        ? readMiniSectorChain(entry.startSector, miniFat, miniStream, miniSectorSize)
        : readSectorChain(entry.startSector, fat, bytes, sectorSize);
      streams.set(entry.name, data.slice(0, entry.streamSize));
    });

  return { streams };
}

function readFibField(wordDocument: Uint8Array, offset: number) {
  if (offset + 4 > wordDocument.length) return 0;
  return readUint32(new DataView(wordDocument.buffer, wordDocument.byteOffset, wordDocument.byteLength), offset);
}

function parseFib(wordDocument: Uint8Array) {
  const view = new DataView(wordDocument.buffer, wordDocument.byteOffset, wordDocument.byteLength);
  const flags = readUint16(view, 10);

  return {
    tableStreamName: flags & 0x0200 ? '1Table' : '0Table',
    ccpText: readFibField(wordDocument, 76),
    fcPlcfBteChpx: readFibField(wordDocument, 250),
    lcbPlcfBteChpx: readFibField(wordDocument, 254),
    fcPlcfBtePapx: readFibField(wordDocument, 258),
    lcbPlcfBtePapx: readFibField(wordDocument, 262),
    fcSttbfFfn: readFibField(wordDocument, 274),
    lcbSttbfFfn: readFibField(wordDocument, 278),
    fcClx: readFibField(wordDocument, 418),
    lcbClx: readFibField(wordDocument, 422),
  };
}

function findPieceTable(clx: Uint8Array) {
  let offset = 0;
  const view = new DataView(clx.buffer, clx.byteOffset, clx.byteLength);

  while (offset < clx.length) {
    const type = clx[offset];

    if (type === 0x02) {
      const length = readUint32(view, offset + 1);
      return clx.slice(offset + 5, offset + 5 + length);
    }

    if (type === 0x01) {
      const length = readUint16(view, offset + 1);
      offset += 3 + length;
      continue;
    }

    offset += 1;
  }

  return undefined;
}

function parsePieces(tableStream: Uint8Array, fib: DocFib) {
  // Piece table 描述正文字符区间与 WordDocument 字节偏移的映射，是读取 DOC 正文的核心索引。
  const clx = tableStream.slice(fib.fcClx, fib.fcClx + fib.lcbClx);
  const pieceTable = findPieceTable(clx);
  if (!pieceTable) return [];

  const pieceCount = Math.floor((pieceTable.length - 4) / 12);
  const view = new DataView(pieceTable.buffer, pieceTable.byteOffset, pieceTable.byteLength);
  const pieces: DocPiece[] = [];

  for (let index = 0; index < pieceCount; index += 1) {
    const charStart = readUint32(view, index * 4);
    const charEnd = readUint32(view, (index + 1) * 4);
    const pcdOffset = (pieceCount + 1) * 4 + index * 8;
    const fcValue = readUint32(view, pcdOffset + 2);
    const compressed = Boolean(fcValue & 0x40000000);
    const fileOffset = compressed ? (fcValue & 0x3fffffff) / 2 : fcValue;

    if (charEnd > charStart && fileOffset >= 0) {
      pieces.push({ charStart, charEnd, fileOffset, compressed });
    }
  }

  return pieces;
}

function quoteFontFamily(value: string | undefined) {
  if (!value) return undefined;
  return value
    .split(',')
    .map((font) => font.trim())
    .filter(Boolean)
    .map((font) => (/^["'].*["']$/.test(font) || /^[a-z-]+$/i.test(font) ? font : `"${font}"`))
    .join(', ');
}

function parseFontTable(tableStream: Uint8Array, fib: DocFib): DocFontTable {
  if (!fib.fcSttbfFfn || !fib.lcbSttbfFfn) return [];
  const data = tableStream.slice(fib.fcSttbfFfn, fib.fcSttbfFfn + fib.lcbSttbfFfn);
  if (data.length < 4) return [];

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const extended = readUint16(view, 0) === 0xffff;
  const count = extended && data.length >= 6 ? readUint16(view, 2) : readUint16(view, 0);
  const fonts: string[] = [];
  let offset = extended ? 6 : 4;

  for (let index = 0; index < count && offset + 1 < data.length; index += 1) {
    const size = data[offset];
    if (!size || offset + size > data.length) break;

    const nameOffset = offset + 40;
    if (nameOffset < offset + size) {
      const rawName = new TextDecoder('utf-16le').decode(data.slice(nameOffset, offset + size));
      const name = rawName.split('\u0000')[0]?.replace(/\uFFFD/g, '').trim();
      if (name) fonts.push(name);
    }

    offset += size;
  }

  return fonts;
}

function plcItemCount(length: number, dataSize: number) {
  return Math.floor((length - 4) / (4 + dataSize));
}

function parsePlcBteChpx(tableStream: Uint8Array, fib: DocFib) {
  if (!fib.fcPlcfBteChpx || !fib.lcbPlcfBteChpx) return [];

  const data = tableStream.slice(fib.fcPlcfBteChpx, fib.fcPlcfBteChpx + fib.lcbPlcfBteChpx);
  const count = plcItemCount(data.length, 4);
  if (count <= 0) return [];

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const pnOffset = (count + 1) * 4;

  return Array.from({ length: count }, (_, index) => ({
    fcStart: readUint32(view, index * 4),
    fcEnd: readUint32(view, (index + 1) * 4),
    pn: readUint32(view, pnOffset + index * 4) & 0x003fffff,
  })).filter((item) => item.fcEnd > item.fcStart);
}

function parsePlcBtePapx(tableStream: Uint8Array, fib: DocFib) {
  if (!fib.fcPlcfBtePapx || !fib.lcbPlcfBtePapx) return [];

  const data = tableStream.slice(fib.fcPlcfBtePapx, fib.fcPlcfBtePapx + fib.lcbPlcfBtePapx);
  const count = plcItemCount(data.length, 4);
  if (count <= 0) return [];

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const pnOffset = (count + 1) * 4;

  return Array.from({ length: count }, (_, index) => ({
    fcStart: readUint32(view, index * 4),
    fcEnd: readUint32(view, (index + 1) * 4),
    pn: readUint32(view, pnOffset + index * 4) & 0x003fffff,
  })).filter((item) => item.fcEnd > item.fcStart);
}

function parsePlcBteTapx(tableStream: Uint8Array, fib: DocFib) {
  const fc = fib.fcPlcfBtePapx + fib.lcbPlcfBtePapx;
  if (!fc) return [];

  const data = tableStream.slice(fc, tableStream.length);
  if (data.length < 8) return [];

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const count = Math.floor((data.length - 4) / 4);
  return Array.from({ length: count }, (_, index) => ({
    fcStart: readUint32(view, index * 4),
    fcEnd: readUint32(view, (index + 1) * 4),
    pn: readUint32(view, (count + 1) * 4 + index * 4) & 0x003fffff,
  })).filter((item) => item.fcEnd > item.fcStart);
}

function mergeTextStyle(base: DocTextStyle | undefined, next: DocTextStyle | undefined): DocTextStyle | undefined {
  if (!base && !next) return undefined;
  return {
    ...base,
    ...next,
  };
}

function firstDefined<T>(...values: Array<T | undefined>) {
  return values.find((value) => value !== undefined);
}

function mergeStyleIntoTextStyle(base: DocTextStyle, override: DocTextStyle | undefined) {
  return {
    ...base,
    ...override,
    fontSize: override?.fontSize ?? base.fontSize,
    fontWeight: override?.fontWeight ?? base.fontWeight,
    fontStyle: override?.fontStyle ?? base.fontStyle,
    textDecoration: override?.textDecoration ?? base.textDecoration,
    color: override?.color ?? base.color,
    backgroundColor: override?.backgroundColor ?? base.backgroundColor,
    textAlign: override?.textAlign ?? base.textAlign,
    lineHeight: override?.lineHeight ?? base.lineHeight,
    fontFamily: override?.fontFamily ?? base.fontFamily,
    indentLeft: override?.indentLeft ?? base.indentLeft,
    indentRight: override?.indentRight ?? base.indentRight,
    firstLineIndent: override?.firstLineIndent ?? base.firstLineIndent,
    spacingBefore: override?.spacingBefore ?? base.spacingBefore,
    spacingAfter: override?.spacingAfter ?? base.spacingAfter,
    paddingTop: override?.paddingTop ?? base.paddingTop,
    paddingRight: override?.paddingRight ?? base.paddingRight,
    paddingBottom: override?.paddingBottom ?? base.paddingBottom,
    paddingLeft: override?.paddingLeft ?? base.paddingLeft,
  };
}

function blockStyleFromTextStyle(style: DocTextStyle | undefined): DocTextStyle | undefined {
  if (!style) return undefined;
  const blockStyle: DocTextStyle = {
    textAlign: style.textAlign,
    lineHeight: style.lineHeight,
    indentLeft: style.indentLeft,
    indentRight: style.indentRight,
    firstLineIndent: style.firstLineIndent,
    spacingBefore: style.spacingBefore,
    spacingAfter: style.spacingAfter,
  };
  const cleaned = Object.fromEntries(
    Object.entries(blockStyle).filter(([, value]) => value !== undefined),
  ) as DocTextStyle;
  return Object.keys(cleaned).length ? cleaned : undefined;
}

function tableCellTextStyle(style: DocTextStyle | undefined): DocTextStyle | undefined {
  if (!style) return undefined;
  const cellTextStyle: DocTextStyle = {
    color: style.color,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    textDecoration: style.textDecoration,
    textAlign: style.textAlign,
    lineHeight: style.lineHeight,
    fontFamily: style.fontFamily,
  };
  const cleaned = Object.fromEntries(
    Object.entries(cellTextStyle).filter(([, value]) => value !== undefined),
  ) as DocTextStyle;
  return Object.keys(cleaned).length ? cleaned : undefined;
}

function mergeTextDecoration(style: DocTextStyle, decoration: string, enabled: boolean) {
  const values = new Set((style.textDecoration ?? '').split(/\s+/).filter(Boolean));
  if (enabled) values.add(decoration);
  else values.delete(decoration);
  style.textDecoration = values.size ? [...values].join(' ') : undefined;
}

function applySprmOperand(style: DocTextStyle, sprm: number, operand: Uint8Array, fonts: DocFontTable = []) {
  const operandView = new DataView(operand.buffer, operand.byteOffset, operand.byteLength);
  const first = operand[0];

  if ((sprm === 0x0835 || sprm === 0x0800) && first !== undefined) {
    style.fontWeight = first ? 700 : 400;
    return;
  }

  if ((sprm === 0x0836 || sprm === 0x0801) && first !== undefined) {
    style.fontStyle = first ? 'italic' : 'normal';
    return;
  }

  if ((sprm === 0x0837 || sprm === 0x0802) && first !== undefined) {
    mergeTextDecoration(style, 'underline', first !== 0);
    return;
  }

  if ((sprm === 0x0838 || sprm === 0x0803) && first !== undefined) {
    mergeTextDecoration(style, 'line-through', first !== 0);
    return;
  }

  if ((sprm === 0x4a43 || sprm === 0x4a4d) && operand.length >= 2) {
    const halfPoints = readInt16(operandView, 0);
    if (halfPoints > 0 && halfPoints < 200) {
      style.fontSize = halfPoints / 2 * (96 / 72);
    }
    return;
  }

  if ((sprm === 0x4a4f || sprm === 0x4a50 || sprm === 0x4a51 || sprm === 0x4a4e) && operand.length >= 2) {
    const font = quoteFontFamily(fonts[readUint16(operandView, 0)]);
    if (font) style.fontFamily = font;
    return;
  }

  if ((sprm === 0x2a42 || sprm === 0x2a24) && first !== undefined) {
    style.color = WORD_ICO_COLORS[first];
    return;
  }

  if ((sprm === 0x2403 || sprm === 0x2461) && first !== undefined) {
    const alignment = ['left', 'center', 'right', 'justify'][first];
    if (alignment) style.textAlign = alignment as DocTextStyle['textAlign'];
    return;
  }

  if ((sprm === 0x840f || sprm === 0x845e) && operand.length >= 2) {
    style.indentLeft = twipToPx(readInt16(operandView, 0));
    return;
  }

  if ((sprm === 0x8411 || sprm === 0x8460) && operand.length >= 2) {
    style.indentRight = twipToPx(readInt16(operandView, 0));
    return;
  }

  if ((sprm === 0x8410 || sprm === 0x845f) && operand.length >= 2) {
    style.firstLineIndent = twipToPx(readInt16(operandView, 0));
    return;
  }

  if ((sprm === 0xA413 || sprm === 0xA416) && operand.length >= 2) {
    const value = readInt16(operandView, 0);
    if (value >= 0) style.spacingBefore = twipToPx(value);
    return;
  }

  if ((sprm === 0xA414 || sprm === 0xA417) && operand.length >= 2) {
    const value = readInt16(operandView, 0);
    if (value >= 0) style.spacingAfter = twipToPx(value);
    return;
  }

  if ((sprm === 0x6412 || sprm === 0x6461) && operand.length >= 4) {
    const line = readInt16(operandView, 0);
    if (line > 0) {
      style.lineHeight = line >= 240 ? line / 240 : twipToPx(line);
    }
  }
}

function applyTableSprmOperand(style: DocTableStyle, sprm: number, operand: Uint8Array) {
  const operandView = new DataView(operand.buffer, operand.byteOffset, operand.byteLength);
  const first = operand[0];

  if ((sprm === 0x5400 || sprm === 0x548a) && operand.length >= 2) {
    const justify = readInt16(operandView, 0);
    style.borderColor = style.borderColor ?? '#cbd5e1';
    if (justify === 1) style.headerBackgroundColor = style.headerBackgroundColor ?? '#eef4ff';
    return;
  }

  if (sprm === 0xD613 && operand.length >= 8) {
    const color = WORD_ICO_COLORS[first ?? 0];
    if (color) style.borderColor = color;
    return;
  }

  if (sprm === 0xD660 && operand.length >= 4) {
    const color = WORD_ICO_COLORS[first ?? 0];
    if (color) style.cellBackgroundColor = color;
    return;
  }

  if (sprm === 0xD609 && operand.length >= 4) {
    style.stripedRowBackgroundColor = '#f8fafc';
  }
}

function sprmOperandSize(sprm: number, bytes: Uint8Array, offset: number) {
  const sizeCode = (sprm >> 13) & 0x7;
  if (sizeCode === 0 || sizeCode === 1) return 1;
  if (sizeCode === 2 || sizeCode === 4 || sizeCode === 5) return 2;
  if (sizeCode === 3) return 4;
  if (sizeCode === 6) {
    const length = bytes[offset] ?? 0;
    return 1 + length;
  }
  if (sizeCode === 7) {
    const length = bytes[offset] ?? 0;
    return 1 + length;
  }
  return 0;
}

function parseGrpprlStyle(bytes: Uint8Array, fonts: DocFontTable = []): DocTextStyle | undefined {
  const style: DocTextStyle = {};
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;

  while (offset + 2 <= bytes.length) {
    const sprm = readUint16(view, offset);
    offset += 2;
    const operandSize = sprmOperandSize(sprm, bytes, offset);
    if (!operandSize || offset + operandSize > bytes.length) break;
    applySprmOperand(style, sprm, bytes.slice(offset, offset + operandSize), fonts);
    offset += operandSize;
  }

  return Object.keys(style).length ? style : undefined;
}

function parseGrpprlTableStyle(bytes: Uint8Array): DocTableStyle | undefined {
  const style: DocTableStyle = {};
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;

  while (offset + 2 <= bytes.length) {
    const sprm = readUint16(view, offset);
    offset += 2;
    const operandSize = sprmOperandSize(sprm, bytes, offset);
    if (!operandSize || offset + operandSize > bytes.length) break;
    applyTableSprmOperand(style, sprm, bytes.slice(offset, offset + operandSize));
    offset += operandSize;
  }

  return Object.keys(style).length ? style : undefined;
}

function parseChpxFkpPage(wordDocument: Uint8Array, pageOffset: number, fonts: DocFontTable): DocCharacterRun[] {
  const page = wordDocument.slice(pageOffset, pageOffset + 512);
  if (page.length < 512) return [];

  const runCount = page[511] ?? 0;
  const view = new DataView(page.buffer, page.byteOffset, page.byteLength);
  const runs: DocCharacterRun[] = [];

  for (let index = 0; index < runCount; index += 1) {
    const fcStart = readUint32(view, index * 4);
    const fcEnd = readUint32(view, (index + 1) * 4);
    const chpxOffset = page[(runCount + 1) * 4 + index];
    if (!chpxOffset || fcEnd <= fcStart) continue;

    const chpxStart = chpxOffset * 2;
    const chpxLength = page[chpxStart] ?? 0;
    const grpprl = page.slice(chpxStart + 1, chpxStart + 1 + chpxLength);
    const style = parseGrpprlStyle(grpprl, fonts);
    if (style) runs.push({ fcStart, fcEnd, style });
  }

  return runs;
}

function parseCharacterRuns(
  wordDocument: Uint8Array,
  tableStream: Uint8Array,
  fib: DocFib,
  fonts: DocFontTable,
): DocCharacterRun[] {
  return parsePlcBteChpx(tableStream, fib).flatMap((entry) =>
    parseChpxFkpPage(wordDocument, entry.pn * 512, fonts).filter(
      (run) => run.fcEnd > entry.fcStart && run.fcStart < entry.fcEnd,
    ),
  );
}

function parsePapxFkpPage(wordDocument: Uint8Array, pageOffset: number): DocParagraphRun[] {
  const page = wordDocument.slice(pageOffset, pageOffset + 512);
  if (page.length < 512) return [];

  const runCount = page[511] ?? 0;
  const view = new DataView(page.buffer, page.byteOffset, page.byteLength);
  const runs: DocParagraphRun[] = [];

  for (let index = 0; index < runCount; index += 1) {
    const fcStart = readUint32(view, index * 4);
    const fcEnd = readUint32(view, (index + 1) * 4);
    const bxOffset = (runCount + 1) * 4 + index * 13;
    const papxOffset = page[bxOffset];
    if (!papxOffset || fcEnd <= fcStart) continue;

    const papxStart = papxOffset * 2;
    const cb = page[papxStart] ?? 0;
    const cbPrime = cb === 0 ? page[papxStart + 1] ?? 0 : cb;
    const papxLength = cb === 0 ? cbPrime * 2 : cb * 2 - 1;
    const grpprlStart = papxStart + (cb === 0 ? 2 : 1);
    const grpprl = page.slice(grpprlStart, grpprlStart + papxLength);
    const style = parseGrpprlStyle(grpprl);
    if (style) runs.push({ fcStart, fcEnd, style });
  }

  return runs;
}

function parseParagraphRuns(wordDocument: Uint8Array, tableStream: Uint8Array, fib: DocFib): DocParagraphRun[] {
  return parsePlcBtePapx(tableStream, fib).flatMap((entry) =>
    parsePapxFkpPage(wordDocument, entry.pn * 512).filter(
      (run) => run.fcEnd > entry.fcStart && run.fcStart < entry.fcEnd,
    ),
  );
}

function parseTableRuns(wordDocument: Uint8Array, tableStream: Uint8Array, fib: DocFib): DocTableRun[] {
  const tableFc = fib.fcPlcfBtePapx + fib.lcbPlcfBtePapx;
  if (!tableFc) return [];

  const data = tableStream.slice(tableFc, tableStream.length);
  if (data.length < 8) return [];

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const count = Math.floor((data.length - 4) / 4);

  return Array.from({ length: count }, (_, index) => {
    const fcStart = readUint32(view, index * 4);
    const fcEnd = readUint32(view, (index + 1) * 4);
    const pn = readUint32(view, (count + 1) * 4 + index * 4) & 0x003fffff;
    const page = wordDocument.slice(pn * 512, pn * 512 + 512);
    if (page.length < 512) return undefined;
    const style = parseGrpprlTableStyle(page.slice(0, page[0] ?? 0));
    if (!style) return undefined;
    return { fcStart, fcEnd, style };
  }).filter((item): item is DocTableRun => Boolean(item));
}

function fileOffsetForPieceChar(piece: DocPiece, charOffset: number) {
  return piece.compressed ? (piece.fileOffset + charOffset) * 2 : piece.fileOffset + charOffset * 2;
}

function pieceCharOffsetForFileOffset(piece: DocPiece, fileOffset: number) {
  return piece.compressed ? fileOffset / 2 - piece.fileOffset : (fileOffset - piece.fileOffset) / 2;
}

function styleForRange(byteStart: number, byteEnd: number, characterRuns: DocCharacterRun[]) {
  return characterRuns
    .filter((run) => run.fcEnd > byteStart && run.fcStart < byteEnd)
    .sort((left, right) => left.fcStart - right.fcStart)
    .reduce<DocTextStyle | undefined>((style, run) => mergeTextStyle(style, run.style), undefined);
}

function paragraphStyleForRange(byteStart: number, byteEnd: number, paragraphRuns: DocParagraphRun[]) {
  return paragraphRuns
    .filter((run) => run.fcEnd > byteStart && run.fcStart < byteEnd)
    .sort((left, right) => left.fcStart - right.fcStart)
    .reduce<DocTextStyle | undefined>((style, run) => mergeTextStyle(style, run.style), undefined);
}

function splitPieceByStyleRuns(piece: DocPiece, characterRuns: DocCharacterRun[], paragraphRuns: DocParagraphRun[]) {
  const charLength = piece.charEnd - piece.charStart;
  const byteStart = piece.compressed ? piece.fileOffset * 2 : piece.fileOffset;
  const byteEnd = byteStart + charLength * 2;
  const boundaries = new Set([0, charLength]);

  [...characterRuns, ...paragraphRuns].forEach((run) => {
    if (run.fcEnd <= byteStart || run.fcStart >= byteEnd) return;
    const start = Math.max(0, Math.floor(pieceCharOffsetForFileOffset(piece, Math.max(run.fcStart, byteStart))));
    const end = Math.min(charLength, Math.ceil(pieceCharOffsetForFileOffset(piece, Math.min(run.fcEnd, byteEnd))));
    boundaries.add(start);
    boundaries.add(end);
  });

  const sorted = [...boundaries].sort((left, right) => left - right);
  return sorted
    .slice(0, -1)
    .map((start, index) => ({ start, end: sorted[index + 1] }))
    .filter((range) => range.end > range.start);
}

function textSegmentsFromPieces(
  wordDocument: Uint8Array,
  pieces: DocPiece[],
  characterRuns: DocCharacterRun[],
  paragraphRuns: DocParagraphRun[],
) {
  return pieces.flatMap((piece) => {
    return splitPieceByStyleRuns(piece, characterRuns, paragraphRuns).map((range) => {
      const scopedPiece: DocPiece = {
        ...piece,
        charStart: piece.charStart + range.start,
        charEnd: piece.charStart + range.end,
        fileOffset: piece.compressed ? piece.fileOffset + range.start : piece.fileOffset + range.start * 2,
      };
      const byteStart = fileOffsetForPieceChar(piece, range.start);
      const byteEnd = fileOffsetForPieceChar(piece, range.end);
      return readPieceSegment(
        wordDocument,
        scopedPiece,
        mergeTextStyle(
          paragraphStyleForRange(byteStart, byteEnd, paragraphRuns),
          styleForRange(byteStart, byteEnd, characterRuns),
        ),
      );
    });
  });
}

function decodeCodePage1252(bytes: Uint8Array) {
  const decoder = typeof TextDecoder !== 'undefined' ? new TextDecoder('windows-1252') : undefined;
  if (decoder) return decoder.decode(bytes);
  return Array.from(bytes, (value) => String.fromCharCode(value)).join('');
}

function bytesToDataUrl(bytes: Uint8Array, mimeType: string) {
  let binary = '';
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return `data:${mimeType};base64,${btoa(binary)}`;
}

function scoreDecodedText(text: string) {
  let score = 0;

  for (const char of text) {
    const code = char.codePointAt(0) ?? 0;
    if (char === '\uFFFD') {
      score -= 12;
    } else if (code === 0) {
      score -= 5;
    } else if (code < 32 && code !== 9 && code !== 10 && code !== 13) {
      score -= 3;
    } else if (/[A-Za-z0-9]/.test(char)) {
      score += 1.2;
    } else if (/\p{Script=Han}/u.test(char)) {
      score += 3.5;
    } else if (/\s/.test(char)) {
      score += 0.2;
    } else if (/[.,:;!?()\-_/|]/.test(char)) {
      score += 0.5;
    } else {
      score += 0.3;
    }
  }

  return score;
}

function decodeCompressedPiece(bytes: Uint8Array) {
  const candidates = ['gb18030', 'gbk', 'utf-8', 'windows-1252'];
  let best = decodeCodePage1252(bytes);
  let bestScore = scoreDecodedText(best);

  for (const encoding of candidates) {
    try {
      const decoder = new TextDecoder(encoding, { fatal: false });
      const decoded = decoder.decode(bytes);
      const score = scoreDecodedText(decoded);
      if (score > bestScore) {
        best = decoded;
        bestScore = score;
      }
    } catch {
      // Ignore unsupported encodings and keep the best available decode.
    }
  }

  return best;
}

function readPieceText(wordDocument: Uint8Array, piece: DocPiece) {
  const charLength = piece.charEnd - piece.charStart;
  if (piece.compressed) {
    return decodeCompressedPiece(wordDocument.slice(piece.fileOffset, piece.fileOffset + charLength));
  }

  const byteLength = charLength * 2;
  const bytes = wordDocument.slice(piece.fileOffset, piece.fileOffset + byteLength);
  return new TextDecoder('utf-16le').decode(bytes);
}

function readPieceSegment(wordDocument: Uint8Array, piece: DocPiece, style?: DocTextStyle): DocTextSegment {
  return {
    text: readPieceText(wordDocument, piece),
    style,
  };
}

function normalizeDocText(text: string) {
  return text
    .replace(/\u0000/g, '')
    .replace(/\u0007/g, '|')
    .replace(/\u000b/g, '\n')
    .replace(/\u000c/g, '\n')
    .replace(/\u000d/g, '\n')
    .replace(/\u0013([^\u0014\u0015]*)(?:\u0014([^\u0015]*))?/g, (_match, instruction = '', result = '') =>
      `${instruction}${result}`.includes('\u0001') ? '\u0001' : '',
    )
    .replace(/\u0015/g, '')
    .replace(/[\u0002-\u0006\u0008\u000e-\u001f]/g, '');
}

function normalizeDocTextSegments(segments: DocTextSegment[], images: DocImage[] = []) {
  let imageIndex = 0;

  return segments.flatMap((segment) => {
    const anchorCount = Array.from(segment.text).filter((char) => char === '\u0001').length;
    const normalizedText =
      anchorCount > 0 && normalizeDocText(segment.text).includes('\u0001')
        ? normalizeDocText(segment.text)
        : anchorCount > 0
          ? `${normalizeDocText(segment.text)}${'\u0001'.repeat(anchorCount)}`
          : normalizeDocText(segment.text);

    return normalizedText
      .split(/(\n|\u0001)/)
      .map((text): DocImageSegment => {
        if (text === '\u0001') {
          const image = images[imageIndex];
          if (image) imageIndex += 1;
          return { text, style: segment.style, image };
        }
        return { text, style: segment.style };
      })
      .filter((item) => item.image || (item.text.length && item.text !== '\u0001'));
  });
}

function normalizeBlockText(text: string) {
  return text.replace(/[ \t]+/g, ' ').trim();
}

function textFromInlines(inlines: DocTextInline[]) {
  return inlines.map((inline) => (inline.type === 'text' ? inline.text : '')).join('');
}

function sameInlineStyle(left?: DocTextStyle, right?: DocTextStyle) {
  return JSON.stringify(left ?? {}) === JSON.stringify(right ?? {});
}

function mergeAdjacentInlines(inlines: DocTextInline[]) {
  const merged: DocTextInline[] = [];

  inlines.forEach((inline) => {
    if (inline.type === 'image') {
      merged.push({ ...inline });
      return;
    }
    if (!inline.text) return;
    const previous = merged[merged.length - 1];
    if (previous?.type === 'text' && sameInlineStyle(previous.style, inline.style)) {
      previous.text += inline.text;
      return;
    }
    merged.push({ ...inline });
  });

  return merged;
}

function trimInlines(inlines: DocTextInline[]) {
  const result = inlines
    .map((inline) => ({ ...inline }))
    .filter((inline) => inline.type === 'image' || (inline.type === 'text' && inline.text.length));

  while (result.length && result[0].type === 'text' && !result[0].text.trim()) {
    result.shift();
  }

  while (result.length) {
    const last = result[result.length - 1];
    if (last.type !== 'text' || last.text.trim()) break;
    result.pop();
  }

  if (result.length) {
    if (result[0].type === 'text') result[0].text = result[0].text.replace(/^\s+/, '');
    const last = result[result.length - 1];
    if (last.type === 'text') last.text = last.text.replace(/\s+$/, '');
  }

  return mergeAdjacentInlines(result);
}

function looksLikeTableRow(line: string) {
  return line.split('|').map((cell) => cell.trim()).filter(Boolean).length >= 2;
}

function splitTableCells(line: DocLine): PendingTableCell[] {
  const cells: PendingTableCell[] = [];
  let current: DocTextInline[] = [];
  const textInlines = (inlines: DocTextInline[]) => inlines.filter((item): item is Extract<DocTextInline, { type: 'text' }> => item.type === 'text');

  line.inlines.forEach((inline) => {
    if (inline.type === 'image') {
      current.push(inline);
      return;
    }
    const parts = inline.text.split('|');

    parts.forEach((part, index) => {
      if (index > 0) {
        const inlines = trimInlines(current);
        if (inlines.length) {
          cells.push({
            text: normalizeBlockText(textFromInlines(inlines)),
            inlines,
            style: dominantStyle(textInlines(inlines).map((item) => ({ text: item.text, style: item.style }))),
          });
        }
        current = [];
      }

      if (part) {
        current.push({ ...inline, text: part });
      }
    });
  });

  const inlines = trimInlines(current);
  if (inlines.length) {
    cells.push({
      text: normalizeBlockText(textFromInlines(inlines)),
      inlines,
      style: dominantStyle(textInlines(inlines).map((item) => ({ text: item.text, style: item.style }))),
    });
  }

  return cells;
}

function sliceLineInlines(line: DocLine, start: number) {
  let offset = 0;
  const result: DocTextInline[] = [];

  line.inlines.forEach((inline) => {
    if (inline.type === 'image') return;
    const inlineStart = offset;
    const inlineEnd = inlineStart + inline.text.length;
    offset = inlineEnd;

    if (inlineEnd <= start) return;

    result.push({
      ...inline,
      text: inline.text.slice(Math.max(0, start - inlineStart)),
    });
  });

  return trimInlines(result);
}

function parseListLine(line: DocLine): ParsedListLine | undefined {
  const orderedMatch = line.text.match(
    /^\s*(?:(?:\(?[0-9A-Za-z]{1,3}\)?[.)\u3001\uff1f])|(?:[\uff08(][\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]{1,3}[\uff09)]))\s+(.+)$/,
  );
  if (orderedMatch?.[1]) {
    const contentStart = orderedMatch[0].length - orderedMatch[1].length;
    return {
      ordered: true,
      text: normalizeBlockText(orderedMatch[1]),
      inlines: sliceLineInlines(line, contentStart),
    };
  }

  const unorderedMatch = line.match(/^\s*(?:[\u2022\u25cf\u25cb\u25a0\u25c6]|[-*])\s+(.+)$/);
  if (unorderedMatch?.[1]) {
    return {
      ordered: false,
      text: normalizeBlockText(unorderedMatch[1]),
    };
  }

  return undefined;
}

function inferParagraphStyle(role: DocParagraphBlock['role'], text: string): DocTextStyle {
  if (role === 'title') {
    return {
      fontSize: 22,
      fontWeight: 700,
      lineHeight: 1.45,
      color: '#111827',
      textAlign: 'left',
      fontFamily: DOC_FONT_FAMILY,
      paddingBottom: 4,
    };
  }

  if (role === 'heading') {
    return {
      fontSize: 16,
      fontWeight: 700,
      lineHeight: 1.65,
      color: '#1f2937',
      textAlign: 'left',
      fontFamily: DOC_FONT_FAMILY,
      paddingTop: 2,
      paddingBottom: 2,
      backgroundColor: text.includes('\u5185\u5bb9\u5757') ? '#f8fafc' : undefined,
    };
  }

  return {
    fontSize: 14,
    fontWeight: 400,
    lineHeight: 1.8,
    color: '#111827',
    textAlign: 'left',
    fontFamily: DOC_FONT_FAMILY,
  };
}

function inferListStyle(ordered: boolean): DocTextStyle {
  return {
    fontSize: 14,
    fontWeight: 400,
    lineHeight: 1.7,
    color: '#111827',
    textAlign: 'left',
    fontFamily: DOC_FONT_FAMILY,
    paddingLeft: ordered ? 2 : 0,
  };
}

function inferTableStyle(): DocTableStyle {
  return {
    headerBackgroundColor: '#eef4ff',
    headerTextColor: '#1d4ed8',
    borderColor: '#cbd5e1',
    cellBackgroundColor: '#ffffff',
    stripedRowBackgroundColor: '#f8fafc',
  };
}

function estimateTableColumns(rows: PendingTableCell[][]) {
  const columnCount = Math.max(...rows.map((row) => row.length), 1);
  const weights = Array.from({ length: columnCount }, (_, columnIndex) =>
    Math.max(
      8,
      ...rows.map((row) => {
        const text = row[columnIndex]?.text ?? '';
        return Array.from(text).reduce((sum, char) => sum + (/[\u4e00-\u9fa5]/.test(char) ? 2 : 1), 0);
      }),
    ),
  );
  const total = weights.reduce((sum, weight) => sum + weight, 0) || 1;
  const availableWidth = DEFAULT_DOC_PAGE.width - DEFAULT_DOC_PAGE.marginLeft - DEFAULT_DOC_PAGE.marginRight;
  return weights.map((weight) => Math.max(64, (weight / total) * availableWidth));
}

function createParagraphBlock(text: string, index: number, inlines?: DocTextInline[], style?: DocTextStyle): DocParagraphBlock {
  const compactLength = text.replace(/\s+/g, '').length;
  const hasImages = Boolean(inlines?.some((inline) => inline.type === 'image'));
  const role =
    index === 0 && compactLength <= 24
      ? 'title'
      : compactLength > 0 && compactLength <= 18 && !hasImages && !/[|:\uff1a]/.test(text) && !/[0-9]{4,}/.test(text)
        ? 'heading'
        : 'body';

  return {
    id: `doc-p-${index + 1}`,
    type: 'paragraph',
    text,
    inlines,
    role,
    style: mergeStyleIntoTextStyle(inferParagraphStyle(role, text), style),
  };
}

function createTableBlock(rows: PendingTableCell[][], index: number): DocTableBlock {
  const tableStyle = inferTableStyle();
  return {
    id: `doc-table-${index + 1}`,
    type: 'table',
    style: tableStyle,
    columns: estimateTableColumns(rows),
    rows: rows.map((row, rowIndex) => ({
      id: `doc-table-${index + 1}-row-${rowIndex + 1}`,
      cells: row.map((cell, cellIndex) => ({
        id: `doc-table-${index + 1}-cell-${rowIndex + 1}-${cellIndex + 1}`,
        text: cell.text,
        inlines: cell.inlines,
        style: {
          color: rowIndex === 0 ? tableStyle.headerTextColor : '#111827',
          backgroundColor:
            rowIndex === 0
              ? tableStyle.headerBackgroundColor
              : rowIndex % 2 === 1
                ? tableStyle.cellBackgroundColor
                : tableStyle.stripedRowBackgroundColor,
          fontSize: rowIndex === 0 ? 13 : 13,
          fontWeight: rowIndex === 0 ? 700 : 400,
          lineHeight: 1.65,
          fontFamily: DOC_FONT_FAMILY,
          paddingTop: 5,
          paddingRight: 8,
          paddingBottom: 5,
          paddingLeft: 8,
          ...tableCellTextStyle(cell.style),
        },
      })),
    })),
  };
}

function createListBlock(items: ParsedListLine[], index: number): DocListBlock {
  const orderedCount = items.filter((item) => item.ordered).length;
  const ordered = orderedCount >= items.length / 2;

  return {
    id: `doc-list-${index + 1}`,
    type: 'list',
    ordered,
    style: inferListStyle(ordered),
    items: items.map((item, itemIndex) => ({
      id: `doc-list-${index + 1}-item-${itemIndex + 1}`,
      text: item.text,
      inlines: item.inlines,
    })),
  };
}

function dominantStyle(segments: DocTextSegment[]) {
  return segments.reduce<DocTextStyle | undefined>((style, segment) => mergeTextStyle(style, segment.style), undefined);
}

function blockHasImage(block: DocBlock) {
  if (block.type === 'paragraph') return Boolean(block.inlines?.some((inline) => inline.type === 'image'));
  if (block.type === 'table') {
    return block.rows.some((row) => row.cells.some((cell) => cell.inlines?.some((inline) => inline.type === 'image')));
  }
  return block.items.some((item) => item.inlines?.some((inline) => inline.type === 'image'));
}

function isImageOnlyParagraph(block: DocBlock) {
  return block.type === 'paragraph' && !block.text.trim() && blockHasImage(block);
}

function isShapeTextParagraph(block: DocBlock) {
  if (block.type !== 'paragraph') return false;
  const text = block.text.replace(/\s+/g, '');
  return (
    !blockHasImage(block) &&
    Boolean(text) &&
    (
      text.includes('\u6dfb\u52a0\u6807\u9898') ||
      text.includes('\u8bf7\u70b9\u51fb\u7f16\u8f91\u6587\u5b57') ||
      text.includes('\u8bf7\u6b64\u5904\u7f16\u8f91\u6587\u5b57')
    )
  );
}

function reorderFloatingShapeTextBlocks(blocks: DocBlock[]) {
  const reordered: DocBlock[] = [];
  let index = 0;

  while (index < blocks.length) {
    if (blocks[index].type !== 'table') {
      reordered.push(blocks[index]);
      index += 1;
      continue;
    }

    const tableBlock = blocks[index];
    let imageEnd = index + 1;
    while (imageEnd < blocks.length && isImageOnlyParagraph(blocks[imageEnd])) {
      imageEnd += 1;
    }

    let textEnd = imageEnd;
    while (textEnd < blocks.length && isShapeTextParagraph(blocks[textEnd])) {
      textEnd += 1;
    }

    if (imageEnd > index + 1 && textEnd > imageEnd) {
      reordered.push(tableBlock, ...blocks.slice(imageEnd, textEnd), ...blocks.slice(index + 1, imageEnd));
      index = textEnd;
      continue;
    }

    reordered.push(tableBlock);
    index += 1;
  }

  return reordered.map((block, blockIndex) => ({ ...block, id: `${block.type === 'table' ? 'doc-table' : block.type === 'list' ? 'doc-list' : 'doc-p'}-${blockIndex + 1}` }));
}

function blocksFromSegments(segments: DocTextSegment[], images: DocImage[] = []): DocBlock[] {
  const blocks: DocBlock[] = [];
  const pendingTableRows: PendingTableCell[][] = [];
  const pendingListItems: ParsedListLine[] = [];
  const normalizedSegments = normalizeDocTextSegments(segments, images);
  const lines: DocLine[] = [];
  let currentLine = '';
  let currentLineInlines: DocTextInline[] = [];
  let currentLineSegments: DocTextSegment[] = [];

  const makeLine = (): DocLine => {
    const text = currentLine;
    return {
      text,
      inlines: mergeAdjacentInlines(trimInlines(currentLineInlines)),
      style: dominantStyle(currentLineSegments),
      match: (pattern) => text.match(pattern),
    };
  };

  normalizedSegments.forEach((segment) => {
    if (segment.text === '\n') {
      lines.push(makeLine());
      currentLine = '';
      currentLineInlines = [];
      currentLineSegments = [];
      return;
    }

    if (segment.image) {
      currentLineInlines.push({ type: 'image', image: segment.image });
      return;
    }

    if (segment.text === '|' && currentLine.endsWith('|')) {
      currentLine = currentLine.slice(0, -1);
      const previousInline = currentLineInlines[currentLineInlines.length - 1];
      if (previousInline?.type === 'text') {
        previousInline.text = previousInline.text.slice(0, -1);
        if (!previousInline.text) currentLineInlines.pop();
      }
      lines.push(makeLine());
      currentLine = '';
      currentLineInlines = [];
      currentLineSegments = [];
      return;
    }

    currentLine += segment.text;
    currentLineInlines.push({ type: 'text', text: segment.text, style: segment.style });
    currentLineSegments.push(segment);
  });

  lines.push(makeLine());

  const flushTable = () => {
    if (!pendingTableRows.length) return;
    if (pendingTableRows.length === 1) {
      const text = pendingTableRows[0].map((cell) => cell.text).join(' ');
      const inlines = pendingTableRows[0].flatMap((cell) => cell.inlines);
      blocks.push(createParagraphBlock(text, blocks.length, inlines));
    } else {
      blocks.push(createTableBlock([...pendingTableRows], blocks.length));
    }
    pendingTableRows.length = 0;
  };

  const flushList = () => {
    if (!pendingListItems.length) return;
    if (pendingListItems.length === 1) {
      blocks.push(createParagraphBlock(pendingListItems[0].text, blocks.length, pendingListItems[0].inlines));
    } else {
      blocks.push(createListBlock([...pendingListItems], blocks.length));
    }
    pendingListItems.length = 0;
  };

  lines.forEach((line) => {
    const textLine = normalizeBlockText(line.text);
    if (!textLine) {
      flushTable();
      flushList();
      if (line.inlines.some((inline) => inline.type === 'image')) {
        blocks.push(createParagraphBlock('', blocks.length, line.inlines, line.style));
      }
      return;
    }

    if (looksLikeTableRow(textLine)) {
      flushList();
      pendingTableRows.push(splitTableCells(line));
      return;
    }

    const listLine = parseListLine(line);
    if (listLine) {
      flushTable();
      if (!listLine.inlines?.length) {
        listLine.inlines = line.inlines;
      }
      pendingListItems.push(listLine);
      return;
    }

    flushTable();
    flushList();
    blocks.push(createParagraphBlock(textLine, blocks.length, line.inlines, line.style));
  });

  flushTable();
  flushList();
  return reorderFloatingShapeTextBlocks(blocks);
}

function blocksFromText(text: string): DocBlock[] {
  return blocksFromSegments([{ text }]);
}

function paragraphsFromBlocks(blocks: DocBlock[]): DocParagraph[] {
  return blocks
    .flatMap((block) => {
      if (block.type === 'paragraph') return [block.text];
      if (block.type === 'list') return block.items.map((item) => item.text);
      return block.rows.map((row) => row.cells.map((cell) => cell.text).join(' '));
    })
    .filter(Boolean)
    .map((text, index) => ({
      id: `doc-summary-p-${index + 1}`,
      text,
    }));
}

function extractImageAt(bytes: Uint8Array, start: number) {
  if (
    bytes[start] === 0x89 &&
    bytes[start + 1] === 0x50 &&
    bytes[start + 2] === 0x4e &&
    bytes[start + 3] === 0x47 &&
    bytes[start + 4] === 0x0d &&
    bytes[start + 5] === 0x0a &&
    bytes[start + 6] === 0x1a &&
    bytes[start + 7] === 0x0a
  ) {
    let offset = start + 8;
    while (offset + 12 <= bytes.length) {
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      const chunkLength = readUint32BE(view, offset);
      const chunkType = String.fromCharCode(bytes[offset + 4], bytes[offset + 5], bytes[offset + 6], bytes[offset + 7]);
      const nextOffset = offset + 12 + chunkLength;
      if (nextOffset > bytes.length) break;
      offset = nextOffset;
      if (chunkType === 'IEND') {
        return {
          mimeType: 'image/png',
          bytes: bytes.slice(start, offset),
        };
      }
    }
  }

  if (bytes[start] === 0xff && bytes[start + 1] === 0xd8 && bytes[start + 2] === 0xff) {
    for (let index = start + 2; index + 1 < bytes.length; index += 1) {
      if (bytes[index] === 0xff && bytes[index + 1] === 0xd9) {
        return {
          mimeType: 'image/jpeg',
          bytes: bytes.slice(start, index + 2),
        };
      }
    }
  }

  return undefined;
}

function readImageSize(bytes: Uint8Array, mimeType: string) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  if (mimeType === 'image/png' && bytes.length >= 24) {
    return {
      width: readUint32BE(view, 16),
      height: readUint32BE(view, 20),
    };
  }

  if (mimeType === 'image/jpeg') {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = bytes[offset + 1];
      const length = readUint16BE(view, offset + 2);
      if (marker >= 0xc0 && marker <= 0xc3 && offset + 8 < bytes.length) {
        return {
          width: readUint16BE(view, offset + 7),
          height: readUint16BE(view, offset + 5),
        };
      }
      offset += Math.max(2, length + 2);
    }
  }

  return {};
}

function textNearBytes(bytes: Uint8Array, start: number, before = 320, after = 80) {
  const slice = bytes.slice(Math.max(0, start - before), Math.min(bytes.length, start + after));
  return Array.from(slice, (value) => (value >= 32 && value <= 126 ? String.fromCharCode(value) : ' ')).join('');
}

function isLikelySameImageObject(left: DocImageCandidate, right: DocImageCandidate) {
  if (left.streamName !== right.streamName) return false;
  if (left.mimeType !== right.mimeType || !left.width || !left.height || !right.width || !right.height) return false;
  if (left.width !== right.width || left.height !== right.height) return false;

  const byteDelta = Math.abs(left.byteLength - right.byteLength);
  const isCloseLength = byteDelta <= 1024 || byteDelta / Math.max(left.byteLength, right.byteLength) <= 0.02;
  const isOfficePreviewPair =
    (left.packagedMedia && right.webExtensionPreview) || (left.webExtensionPreview && right.packagedMedia);
  const isNearAlternatePreview = Math.abs(left.offset - right.offset) <= 120000;

  return isCloseLength && isOfficePreviewPair && isNearAlternatePreview;
}

function chooseBetterImageCandidate(left: DocImageCandidate, right: DocImageCandidate) {
  if (left.packagedMedia !== right.packagedMedia) return left.packagedMedia ? left : right;
  if (left.byteLength !== right.byteLength) return left.byteLength > right.byteLength ? left : right;
  return left.offset <= right.offset ? left : right;
}

function normalizeImageCandidates(candidates: DocImageCandidate[]) {
  const normalized: DocImageCandidate[] = [];

  candidates.forEach((candidate) => {
    const duplicateIndex = normalized.findIndex((image) => isLikelySameImageObject(image, candidate));
    if (duplicateIndex === -1) {
      normalized.push(candidate);
      return;
    }

    normalized[duplicateIndex] = chooseBetterImageCandidate(normalized[duplicateIndex], candidate);
  });

  return normalized
    .sort((left, right) => left.offset - right.offset)
    .map(({ byteLength, packagedMedia, webExtensionPreview, streamName, ...image }, index) => ({
      ...image,
      id: `doc-image-${index + 1}`,
    }));
}

function extractDocImagesFromStream(bytes: Uint8Array, streamName: string) {
  const candidates: DocImageCandidate[] = [];
  const seen = new Set<string>();
  const signatures = [
    { mimeType: 'image/png', header: [0x89, 0x50, 0x4e, 0x47] },
    { mimeType: 'image/jpeg', header: [0xff, 0xd8, 0xff] },
  ];

  for (let index = 0; index < bytes.length - 4; index += 1) {
    const signature = signatures.find(({ header }) =>
      header.every((value, headerIndex) => bytes[index + headerIndex] === value),
    );
    if (!signature) continue;

    const extracted = extractImageAt(bytes, index);
    if (!extracted || extracted.bytes.length < 128) continue;

    const head = Array.from(extracted.bytes.slice(0, 16)).join(',');
    const tail = Array.from(extracted.bytes.slice(Math.max(0, extracted.bytes.length - 16))).join(',');
    const key = `${extracted.mimeType}:${extracted.bytes.length}:${head}:${tail}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const context = textNearBytes(bytes, index);

    candidates.push({
      id: '',
      mimeType: extracted.mimeType,
      src: bytesToDataUrl(extracted.bytes, extracted.mimeType),
      offset: index,
      byteLength: extracted.bytes.length,
      packagedMedia: /drs\/media|drs\\media/.test(context),
      webExtensionPreview: /drs\/webExtensions|drs\\webExtensions/.test(context),
      streamName,
      ...readImageSize(extracted.bytes, extracted.mimeType),
    });
  }

  return candidates;
}

function extractDocImages(cfb: CfbFile) {
  const candidates = Array.from(cfb.streams.entries()).flatMap(([streamName, stream]) =>
    extractDocImagesFromStream(stream, streamName),
  );

  return normalizeImageCandidates(candidates);
}

function parsePlainLikeDoc(bytes: Uint8Array, fileName: string, warnings: string[]): DocDocument {
  const fullText = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  const isRtf = fullText.trimStart().startsWith('{\\rtf');
  const text = isRtf
    ? fullText
        .replace(/\\'[0-9a-f]{2}/gi, '')
        .replace(/\\[a-z]+-?\d* ?/gi, '')
        .replace(/[{}]/g, '')
    : fullText.replace(/<[^>]+>/g, ' ');

  warnings.push(
    isRtf
      ? '\u68c0\u6d4b\u5230 RTF \u5185\u5bb9\uff0c\u5df2\u6309\u7eaf\u6587\u672c\u964d\u7ea7\u9884\u89c8\u3002'
      : '\u68c0\u6d4b\u5230\u975e OLE DOC \u5185\u5bb9\uff0c\u5df2\u6309\u7eaf\u6587\u672c\u964d\u7ea7\u9884\u89c8\u3002',
  );
  return buildDocDocument(fileName, blocksFromText(text), warnings);
}

function buildDocDocument(fileName: string, blocks: DocBlock[], warnings: string[]): DocDocument {
  const paragraphs = paragraphsFromBlocks(blocks);
  const title = paragraphs.find((paragraph) => paragraph.text)?.text ?? (fileName || 'DOC \u6587\u6863');
  const images = [] as DocImage[];

  return {
    title,
    page: DEFAULT_DOC_PAGE,
    blocks,
    paragraphs,
    images,
    warnings,
  };
}

export async function parseDoc(file: File): Promise<DocDocument> {
  // 非 OLE 文件按纯文本降级处理；OLE DOC 则解析 CFB、FIB、piece table 和样式 run。
  const bytes = await readBytes(file);
  const warnings: string[] = [];

  if (!isOleDoc(bytes)) {
    return parsePlainLikeDoc(bytes, file.name, warnings);
  }

  const cfb = parseCfb(bytes);
  const wordDocument = cfb.streams.get('WordDocument');

  if (!wordDocument) {
    throw new Error('DOC \u6587\u4ef6\u7f3a\u5c11 WordDocument \u6570\u636e\u6d41');
  }

  const fib = parseFib(wordDocument);
  const tableStream = cfb.streams.get(fib.tableStreamName);

  if (!tableStream) {
    throw new Error(`DOC \u6587\u4ef6\u7f3a\u5c11 ${fib.tableStreamName} \u6570\u636e\u6d41`);
  }

  const pieces = parsePieces(tableStream, fib);
  if (!pieces.length) {
    throw new Error('\u6682\u672a\u80fd\u8bc6\u522b\u8be5 DOC \u6587\u4ef6\u7684\u6b63\u6587\u7247\u6bb5\u8868');
  }

  const fonts = parseFontTable(tableStream, fib);
  const characterRuns = parseCharacterRuns(wordDocument, tableStream, fib, fonts);
  const paragraphRuns = parseParagraphRuns(wordDocument, tableStream, fib);
  const images = extractDocImages(cfb);
  const segments = textSegmentsFromPieces(wordDocument, pieces, characterRuns, paragraphRuns);
  const blocks = blocksFromSegments(segments, images);

  if (!blocks.length) {
    throw new Error('\u8be5 DOC \u6587\u4ef6\u672a\u89e3\u6790\u5230\u53ef\u9884\u89c8\u6b63\u6587');
  }

  warnings.push(
    images.length
      ? '\u5f53\u524d\u4e3a\u7eaf\u524d\u7aef DOC \u964d\u7ea7\u9884\u89c8\uff0c\u5df2\u63d0\u53d6\u5230\u6587\u6863\u5185\u56fe\u7247\uff0c\u4f46\u6682\u672a\u6062\u590d\u7cbe\u786e\u951a\u70b9\u3001\u590d\u6742\u6837\u5f0f\u548c\u5206\u9875\u3002'
      : '\u5f53\u524d\u4e3a\u7eaf\u524d\u7aef DOC \u964d\u7ea7\u9884\u89c8\uff0c\u6682\u4e0d\u8fd8\u539f\u590d\u6742\u6837\u5f0f\u3001\u56fe\u7247\u951a\u70b9\u548c\u5206\u9875\u3002',
  );
  const document = buildDocDocument(file.name, blocks, warnings);
  document.images = images;
  return document;
}

