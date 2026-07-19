import JSZip from 'jszip';

export type OfficeZipInput = File | Blob | ArrayBuffer | Uint8Array;
export type OfficeEntryMap = Map<string, string | Uint8Array>;

const OFFICE_ENTRY_READ_CONCURRENCY = 4;

type OfficeEntryResult = readonly [path: string, data: string | Uint8Array];

/**
 * 读取 Office ZIP 包中的全部文件，并限制同时解压的条目数量以降低瞬时资源峰值。
 */
export async function loadOfficeEntries(
  file: OfficeZipInput,
): Promise<OfficeEntryMap> {
  const source =
    typeof Blob !== 'undefined' && file instanceof Blob
      ? await file.arrayBuffer()
      : file;
  const zip = await JSZip.loadAsync(source);
  const archiveEntries = Object.entries(zip.files).filter(
    ([, entry]) => !entry.dir,
  );
  const results = new Array<OfficeEntryResult>(archiveEntries.length);
  let nextIndex = 0;

  async function readNextEntry(): Promise<void> {
    while (nextIndex < archiveEntries.length) {
      const entryIndex = nextIndex;
      nextIndex += 1;
      const [path, entry] = archiveEntries[entryIndex];

      try {
        const isXml = /\.xml$/i.test(path) || /\.rels$/i.test(path);
        const data = await entry.async(isXml ? 'text' : 'uint8array');
        results[entryIndex] = [path, data];
      } catch (error) {
        const message = error instanceof Error ? error.message : '未知错误';
        throw new Error(`Office 包条目解压失败（${path}）：${message}`);
      }
    }
  }

  const workerCount = Math.min(
    OFFICE_ENTRY_READ_CONCURRENCY,
    archiveEntries.length,
  );
  await Promise.all(Array.from({ length: workerCount }, () => readNextEntry()));

  return new Map(results);
}

export function readXml(entries: OfficeEntryMap, path: string) {
  const value = entries.get(path);
  return typeof value === 'string' ? value : '';
}

export function readBinary(entries: OfficeEntryMap, path: string) {
  const value = entries.get(path);
  return value instanceof Uint8Array ? value : undefined;
}
