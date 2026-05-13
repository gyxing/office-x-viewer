import JSZip from 'jszip';

export type OfficeZipInput = File | Blob | ArrayBuffer | Uint8Array;
export type OfficeEntryMap = Map<string, string | Uint8Array>;

export async function loadOfficeEntries(file: OfficeZipInput): Promise<OfficeEntryMap> {
  const source = typeof Blob !== 'undefined' && file instanceof Blob ? await file.arrayBuffer() : file;
  const zip = await JSZip.loadAsync(source);
  const entries: OfficeEntryMap = new Map();

  const reads = Object.entries(zip.files).map(async ([path, entry]) => {
    if (entry.dir) {
      return;
    }

    const isXml = /\.xml$/i.test(path) || /\.rels$/i.test(path);
    const data = await entry.async(isXml ? 'text' : 'uint8array');
    entries.set(path, data);
  });

  await Promise.all(reads);
  return entries;
}

export function readXml(entries: OfficeEntryMap, path: string) {
  const value = entries.get(path);
  return typeof value === 'string' ? value : '';
}

export function readBinary(entries: OfficeEntryMap, path: string) {
  const value = entries.get(path);
  return value instanceof Uint8Array ? value : undefined;
}
