import { CfbParseError } from './CfbParseError';
import {
  CFB_DIRECTORY_ENTRY_SIZE,
  CFB_HEADER_SIZE,
  CFB_MINI_SECTOR_SIZE,
  CFB_SIGNATURE,
  DIFAT_SECTOR,
  END_OF_CHAIN,
  FAT_SECTOR,
  FREE_SECTOR,
  MINI_STREAM_CUTOFF_SIZE,
  NO_STREAM,
} from './constants';
import type {
  CfbDirectoryEntry,
  CfbFile,
  CfbObjectType,
  CfbReadOptions,
} from './types';

type CfbHeader = {
  majorVersion: number;
  sectorSize: number;
  sectorCount: number;
  fatSectorCount: number;
  directoryStartSector: number;
  directorySectorCount: number;
  miniFatStartSector: number;
  miniFatSectorCount: number;
  difatStartSector: number;
  difatSectorCount: number;
};

type RawDirectoryEntry = Omit<CfbDirectoryEntry, 'path'>;

function readUint16(view: DataView, offset: number) {
  return view.getUint16(offset, true);
}

function readUint32(view: DataView, offset: number) {
  return view.getUint32(offset, true);
}

function isSpecialSector(value: number) {
  return (
    value === FREE_SECTOR ||
    value === END_OF_CHAIN ||
    value === FAT_SECTOR ||
    value === DIFAT_SECTOR
  );
}

function validateSectorIndex(sector: number, header: CfbHeader) {
  if (!Number.isInteger(sector) || sector < 0 || sector >= header.sectorCount) {
    throw new CfbParseError(
      'SECTOR_OUT_OF_RANGE',
      `CFB 扇区 ${sector} 超出有效范围`,
      { sector },
    );
  }
}

function getSector(bytes: Uint8Array, sector: number, header: CfbHeader) {
  validateSectorIndex(sector, header);
  const offset = (sector + 1) * header.sectorSize;
  const end = offset + header.sectorSize;
  if (offset < 0 || end > bytes.length) {
    throw new CfbParseError(
      'SECTOR_OUT_OF_RANGE',
      `CFB 扇区 ${sector} 的字节范围无效`,
      { sector },
    );
  }
  return bytes.subarray(offset, end);
}

function parseHeader(bytes: Uint8Array): CfbHeader {
  if (bytes.length < CFB_HEADER_SIZE) {
    throw new CfbParseError('INVALID_HEADER', 'CFB Header 长度不足');
  }
  if (!CFB_SIGNATURE.every((value, index) => bytes[index] === value)) {
    throw new CfbParseError('INVALID_SIGNATURE', '不是有效的 CFB 文件');
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const majorVersion = readUint16(view, 26);
  const byteOrder = readUint16(view, 28);
  const sectorShift = readUint16(view, 30);
  const miniSectorShift = readUint16(view, 32);
  const miniStreamCutoff = readUint32(view, 56);
  const expectedSectorShift = majorVersion === 3 ? 9 : 12;

  if (
    (majorVersion !== 3 && majorVersion !== 4) ||
    byteOrder !== 0xfffe ||
    sectorShift !== expectedSectorShift ||
    miniSectorShift !== 6 ||
    miniStreamCutoff !== MINI_STREAM_CUTOFF_SIZE
  ) {
    throw new CfbParseError('INVALID_HEADER', 'CFB Header 固定字段无效');
  }

  const sectorSize = 2 ** sectorShift;
  if (bytes.length < sectorSize || bytes.length % sectorSize !== 0) {
    throw new CfbParseError(
      'INVALID_HEADER',
      'CFB 文件长度与声明的扇区大小不一致',
    );
  }

  const header: CfbHeader = {
    majorVersion,
    sectorSize,
    sectorCount: bytes.length / sectorSize - 1,
    directorySectorCount: readUint32(view, 40),
    fatSectorCount: readUint32(view, 44),
    directoryStartSector: readUint32(view, 48),
    miniFatStartSector: readUint32(view, 60),
    miniFatSectorCount: readUint32(view, 64),
    difatStartSector: readUint32(view, 68),
    difatSectorCount: readUint32(view, 72),
  };

  if (
    header.sectorCount < 1 ||
    (majorVersion === 3 && header.directorySectorCount !== 0) ||
    header.fatSectorCount > header.sectorCount ||
    header.miniFatSectorCount > header.sectorCount ||
    header.difatSectorCount > header.sectorCount
  ) {
    throw new CfbParseError('INVALID_HEADER', 'CFB Header 扇区计数无效');
  }
  validateSectorIndex(header.directoryStartSector, header);
  return header;
}

function concatChunks(chunks: Uint8Array[]) {
  const length = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

async function readDifat(
  bytes: Uint8Array,
  header: CfbHeader,
  options: CfbReadOptions,
) {
  const headerView = new DataView(
    bytes.buffer,
    bytes.byteOffset,
    CFB_HEADER_SIZE,
  );
  const difat: number[] = [];
  for (let offset = 76; offset < CFB_HEADER_SIZE; offset += 4) {
    const sector = readUint32(headerView, offset);
    if (sector !== FREE_SECTOR) difat.push(sector);
  }

  const visited = new Set<number>();
  let sector = header.difatStartSector;
  const entriesPerSector = header.sectorSize / 4 - 1;
  for (let index = 0; index < header.difatSectorCount; index += 1) {
    if (sector === END_OF_CHAIN) {
      throw new CfbParseError('CHAIN_TRUNCATED', 'CFB DIFAT 链提前结束');
    }
    validateSectorIndex(sector, header);
    if (visited.has(sector)) {
      throw new CfbParseError('CHAIN_CYCLE', `CFB DIFAT 扇区 ${sector} 成环`, {
        sector,
      });
    }
    visited.add(sector);
    const data = getSector(bytes, sector, header);
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    for (let entry = 0; entry < entriesPerSector; entry += 1) {
      const fatSector = readUint32(view, entry * 4);
      if (fatSector !== FREE_SECTOR) difat.push(fatSector);
    }
    sector = readUint32(view, header.sectorSize - 4);
    await options.yieldIfNeeded?.();
  }

  if (
    (header.difatSectorCount === 0 &&
      header.difatStartSector !== END_OF_CHAIN) ||
    (header.difatSectorCount > 0 && sector !== END_OF_CHAIN)
  ) {
    throw new CfbParseError('INVALID_HEADER', 'CFB DIFAT 链计数不一致');
  }
  if (difat.length < header.fatSectorCount) {
    throw new CfbParseError('CHAIN_TRUNCATED', 'CFB FAT 扇区列表不完整');
  }
  return difat.slice(0, header.fatSectorCount);
}

async function readFat(
  bytes: Uint8Array,
  header: CfbHeader,
  difat: number[],
  options: CfbReadOptions,
) {
  const fat: number[] = [];
  const visited = new Set<number>();
  for (const sector of difat) {
    validateSectorIndex(sector, header);
    if (visited.has(sector)) {
      throw new CfbParseError('CHAIN_CYCLE', `CFB FAT 扇区 ${sector} 重复`, {
        sector,
      });
    }
    visited.add(sector);
    const data = getSector(bytes, sector, header);
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    for (let offset = 0; offset < data.length; offset += 4) {
      fat.push(readUint32(view, offset));
    }
    await options.yieldIfNeeded?.();
  }
  return fat;
}

async function readSectorChain(
  startSector: number,
  fat: number[],
  bytes: Uint8Array,
  header: CfbHeader,
  options: CfbReadOptions,
  expectedSectors?: number,
) {
  if (startSector === END_OF_CHAIN && (expectedSectors ?? 0) === 0) {
    return new Uint8Array();
  }
  if (isSpecialSector(startSector)) {
    throw new CfbParseError('CHAIN_TRUNCATED', 'CFB 扇区链起点无效', {
      sector: startSector,
    });
  }

  const chunks: Uint8Array[] = [];
  const visited = new Set<number>();
  let sector = startSector;
  while (sector !== END_OF_CHAIN) {
    validateSectorIndex(sector, header);
    if (visited.has(sector)) {
      throw new CfbParseError('CHAIN_CYCLE', `CFB 扇区链在 ${sector} 成环`, {
        sector,
      });
    }
    if (sector >= fat.length) {
      throw new CfbParseError(
        'CHAIN_TRUNCATED',
        `CFB FAT 缺少扇区 ${sector} 的链项`,
        { sector },
      );
    }
    visited.add(sector);
    chunks.push(getSector(bytes, sector, header));
    const next = fat[sector];
    if (next !== END_OF_CHAIN && isSpecialSector(next)) {
      throw new CfbParseError(
        'CHAIN_TRUNCATED',
        `CFB 扇区 ${sector} 指向无效标记`,
        { sector },
      );
    }
    sector = next;
    await options.yieldIfNeeded?.();
  }

  if (expectedSectors !== undefined && chunks.length !== expectedSectors) {
    throw new CfbParseError(
      'CHAIN_TRUNCATED',
      `CFB 扇区链长度 ${chunks.length} 与声明值 ${expectedSectors} 不一致`,
    );
  }
  return concatChunks(chunks);
}

function decodeDirectoryName(bytes: Uint8Array, nameLength: number) {
  if (nameLength < 2 || nameLength > 64 || nameLength % 2 !== 0) {
    throw new CfbParseError('DIRECTORY_CORRUPTED', 'CFB 目录项名称长度无效');
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let result = '';
  for (let offset = 0; offset < nameLength - 2; offset += 2) {
    result += String.fromCharCode(readUint16(view, offset));
  }
  return result;
}

function objectTypeFromValue(value: number): CfbObjectType | undefined {
  if (value === 1) return 'storage';
  if (value === 2) return 'stream';
  if (value === 5) return 'root';
  return undefined;
}

async function parseDirectoryEntries(
  directoryStream: Uint8Array,
  header: CfbHeader,
  options: CfbReadOptions,
) {
  if (directoryStream.length % CFB_DIRECTORY_ENTRY_SIZE !== 0) {
    throw new CfbParseError(
      'DIRECTORY_CORRUPTED',
      'CFB 目录流未按 128 字节对齐',
    );
  }

  const entries: Array<RawDirectoryEntry | undefined> = [];
  for (
    let offset = 0, id = 0;
    offset < directoryStream.length;
    offset += CFB_DIRECTORY_ENTRY_SIZE, id += 1
  ) {
    const data = directoryStream.subarray(
      offset,
      offset + CFB_DIRECTORY_ENTRY_SIZE,
    );
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const rawObjectType = data[66];
    if (rawObjectType === 0) {
      entries.push(undefined);
      continue;
    }
    const objectType = objectTypeFromValue(rawObjectType);
    if (!objectType) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 目录项 ${id} 的对象类型无效`,
        { directoryId: id },
      );
    }
    const lowSize = readUint32(view, 120);
    const highSize = header.majorVersion === 3 ? 0 : readUint32(view, 124);
    const streamSize = highSize * 0x100000000 + lowSize;
    if (!Number.isSafeInteger(streamSize)) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 目录项 ${id} 的流长度超出安全范围`,
        { directoryId: id },
      );
    }
    entries.push({
      id,
      name: decodeDirectoryName(data.subarray(0, 64), readUint16(view, 64)),
      objectType,
      startSector: readUint32(view, 116),
      streamSize,
      leftSiblingId: readUint32(view, 68),
      rightSiblingId: readUint32(view, 72),
      childId: readUint32(view, 76),
    });
    await options.yieldIfNeeded?.();
  }
  return entries;
}

function getDirectoryEntry(
  entries: Array<RawDirectoryEntry | undefined>,
  id: number,
) {
  if (id === NO_STREAM) return undefined;
  const entry = entries[id];
  if (!entry) {
    throw new CfbParseError('DIRECTORY_CORRUPTED', `CFB 目录项 ${id} 不存在`, {
      directoryId: id,
    });
  }
  return entry;
}

function assignDirectoryPaths(
  rawEntries: Array<RawDirectoryEntry | undefined>,
) {
  const root = rawEntries.find((entry) => entry?.objectType === 'root');
  if (!root) {
    throw new CfbParseError('DIRECTORY_CORRUPTED', 'CFB 目录缺少 Root Entry');
  }

  const result: CfbDirectoryEntry[] = [{ ...root, path: '/' }];
  const visited = new Set<number>([root.id]);
  const stack: Array<{ id: number; parentPath: string }> = [];
  if (root.childId !== NO_STREAM) {
    stack.push({ id: root.childId, parentPath: '' });
  }

  while (stack.length) {
    const current = stack.pop()!;
    const entry = getDirectoryEntry(rawEntries, current.id)!;
    if (entry.objectType === 'root' || visited.has(entry.id)) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 目录树在目录项 ${entry.id} 成环`,
        { directoryId: entry.id },
      );
    }
    visited.add(entry.id);
    const path = `${current.parentPath}/${entry.name}`;
    result.push({ ...entry, path });

    if (entry.rightSiblingId !== NO_STREAM) {
      stack.push({
        id: entry.rightSiblingId,
        parentPath: current.parentPath,
      });
    }
    if (entry.leftSiblingId !== NO_STREAM) {
      stack.push({ id: entry.leftSiblingId, parentPath: current.parentPath });
    }
    if (entry.objectType === 'storage' && entry.childId !== NO_STREAM) {
      stack.push({ id: entry.childId, parentPath: path });
    } else if (entry.objectType === 'stream' && entry.childId !== NO_STREAM) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 流目录项 ${entry.id} 不应包含子项`,
        { directoryId: entry.id },
      );
    }
  }

  const unreachable = rawEntries.filter((entry): entry is RawDirectoryEntry =>
    Boolean(entry && !visited.has(entry.id)),
  );
  if (unreachable.length) {
    throw new CfbParseError(
      'DIRECTORY_CORRUPTED',
      `CFB 目录存在不可达目录项 ${unreachable[0].id}`,
      { directoryId: unreachable[0].id },
    );
  }
  return result.sort((left, right) => left.id - right.id);
}

async function readMiniFat(
  bytes: Uint8Array,
  header: CfbHeader,
  fat: number[],
  options: CfbReadOptions,
) {
  if (header.miniFatSectorCount === 0) return [];
  const data = await readSectorChain(
    header.miniFatStartSector,
    fat,
    bytes,
    header,
    options,
    header.miniFatSectorCount,
  );
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const miniFat: number[] = [];
  for (let offset = 0; offset < data.length; offset += 4) {
    miniFat.push(readUint32(view, offset));
  }
  return miniFat;
}

async function readMiniSectorChain(
  entry: CfbDirectoryEntry,
  miniFat: number[],
  miniStream: Uint8Array,
  options: CfbReadOptions,
) {
  if (entry.streamSize === 0) return new Uint8Array();
  const chunks: Uint8Array[] = [];
  const visited = new Set<number>();
  const expectedSectors = Math.ceil(entry.streamSize / CFB_MINI_SECTOR_SIZE);
  let sector = entry.startSector;

  while (sector !== END_OF_CHAIN) {
    if (
      !Number.isInteger(sector) ||
      sector < 0 ||
      sector >= miniFat.length ||
      (sector + 1) * CFB_MINI_SECTOR_SIZE > miniStream.length
    ) {
      throw new CfbParseError(
        'SECTOR_OUT_OF_RANGE',
        `CFB 小流目录项 ${entry.id} 的扇区 ${sector} 无效`,
        { sector, directoryId: entry.id },
      );
    }
    if (visited.has(sector)) {
      throw new CfbParseError(
        'CHAIN_CYCLE',
        `CFB 小流目录项 ${entry.id} 的扇区链成环`,
        { sector, directoryId: entry.id },
      );
    }
    visited.add(sector);
    const offset = sector * CFB_MINI_SECTOR_SIZE;
    chunks.push(miniStream.subarray(offset, offset + CFB_MINI_SECTOR_SIZE));
    const next = miniFat[sector];
    if (next !== END_OF_CHAIN && isSpecialSector(next)) {
      throw new CfbParseError(
        'CHAIN_TRUNCATED',
        `CFB 小流目录项 ${entry.id} 指向无效标记`,
        { sector, directoryId: entry.id },
      );
    }
    sector = next;
    await options.yieldIfNeeded?.();
  }
  if (chunks.length !== expectedSectors) {
    throw new CfbParseError(
      'CHAIN_TRUNCATED',
      `CFB 小流目录项 ${entry.id} 的链长度与流大小不一致`,
      { directoryId: entry.id },
    );
  }
  return concatChunks(chunks).subarray(0, entry.streamSize);
}

function createCfbFile(
  entries: CfbDirectoryEntry[],
  streams: Map<string, Uint8Array>,
): CfbFile {
  const leafEntries = entries.filter((entry) => entry.objectType === 'stream');
  return {
    entries,
    streams,
    getStream: (...names) => {
      for (const name of names) {
        const exactPath = name.startsWith('/') ? name : `/${name}`;
        const exact = streams.get(name) ?? streams.get(exactPath);
        if (exact) return exact;
        const matches = leafEntries.filter(
          (entry) => entry.name.toLowerCase() === name.toLowerCase(),
        );
        if (matches.length === 1) return streams.get(matches[0].path);
      }
      return undefined;
    },
    hasEntry: (name) => {
      const normalized = name.toLowerCase();
      return entries.some(
        (entry) =>
          entry.path.toLowerCase() === normalized ||
          entry.name.toLowerCase() === normalized,
      );
    },
  };
}

/** 解析 CFB 容器，并返回可按完整路径或唯一流名称读取的数据流。 */
export async function parseCfb(
  input: ArrayBuffer | Uint8Array,
  options: CfbReadOptions = {},
): Promise<CfbFile> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const header = parseHeader(bytes);
  const difat = await readDifat(bytes, header, options);
  const fat = await readFat(bytes, header, difat, options);
  const directoryStream = await readSectorChain(
    header.directoryStartSector,
    fat,
    bytes,
    header,
    options,
    header.majorVersion === 4 ? header.directorySectorCount : undefined,
  );
  const rawEntries = await parseDirectoryEntries(
    directoryStream,
    header,
    options,
  );
  const entries = assignDirectoryPaths(rawEntries);
  const root = entries.find((entry) => entry.objectType === 'root')!;
  const rootSectorCount = Math.ceil(root.streamSize / header.sectorSize);
  const miniStream =
    root.streamSize > 0
      ? (
          await readSectorChain(
            root.startSector,
            fat,
            bytes,
            header,
            options,
            rootSectorCount,
          )
        ).subarray(0, root.streamSize)
      : new Uint8Array();
  const miniFat = await readMiniFat(bytes, header, fat, options);
  const streams = new Map<string, Uint8Array>();

  for (const entry of entries) {
    if (entry.objectType !== 'stream') continue;
    if (streams.has(entry.path)) {
      throw new CfbParseError(
        'DIRECTORY_CORRUPTED',
        `CFB 目录存在重复路径 ${entry.path}`,
        { directoryId: entry.id },
      );
    }
    let data: Uint8Array;
    if (entry.streamSize === 0) {
      data = new Uint8Array();
    } else if (entry.streamSize < MINI_STREAM_CUTOFF_SIZE) {
      data = await readMiniSectorChain(entry, miniFat, miniStream, options);
    } else {
      const sectorCount = Math.ceil(entry.streamSize / header.sectorSize);
      data = (
        await readSectorChain(
          entry.startSector,
          fat,
          bytes,
          header,
          options,
          sectorCount,
        )
      ).subarray(0, entry.streamSize);
    }
    streams.set(entry.path, data);
    await options.yieldIfNeeded?.();
  }
  return createCfbFile(entries, streams);
}
