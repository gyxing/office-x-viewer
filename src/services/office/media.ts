import type { OfficeEntryMap } from './archive';
import { readBinary } from './archive';

export type MediaStore = {
  byPath: Record<string, string>;
  byName: Record<string, string>;
};

export type OfficeRelationship = {
  id: string;
  target: string;
  type?: string;
};

export type OfficeRelationshipMap = Record<string, Record<string, OfficeRelationship>>;

function bytesToBase64(bytes: Uint8Array) {
  let binary = '';
  for (let index = 0; index < bytes.length; index += 1) {
    binary += String.fromCharCode(bytes[index]);
  }
  return btoa(binary);
}

export function bytesToDataUrl(bytes: Uint8Array, contentType = 'image/png') {
  return `data:${contentType};base64,${bytesToBase64(bytes)}`;
}

export function imageMimeType(path: string) {
  const lower = path.toLowerCase();
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  if (lower.endsWith('.webp')) return 'image/webp';
  return 'image/png';
}

export function createMediaStore() {
  const store: MediaStore = {
    byPath: {},
    byName: {},
  };

  function register(path: string, bytes: Uint8Array, contentType = imageMimeType(path)) {
    const dataUrl = bytesToDataUrl(bytes, contentType);
    const fileName = path.split('/').pop() ?? path;
    store.byPath[path] = dataUrl;
    store.byName[fileName] = dataUrl;
    return dataUrl;
  }

  function resolve(pathOrName?: string) {
    if (!pathOrName) {
      return undefined;
    }
    return store.byPath[pathOrName] ?? store.byName[pathOrName];
  }

  return { store, register, resolve };
}

export function normalizeRelationshipTarget(relsPath: string, target: string) {
  if (/^[a-z]+:/i.test(target)) {
    return target;
  }

  const baseDir = relsPath
    .replace(/\/_rels\/[^/]+\.rels$/, '')
    .replace(/\/[^/]+\.rels$/, '');
  const parts = `${baseDir}/${target}`.split('/');
  const normalized: string[] = [];
  parts.forEach((part) => {
    if (!part || part === '.') return;
    if (part === '..') {
      normalized.pop();
      return;
    }
    normalized.push(part);
  });
  return normalized.join('/');
}

export function resolvePackageMediaRef(
  target: string | undefined,
  mediaByPath: Record<string, string>,
  mediaByName: Record<string, string>,
  rootDir: string,
) {
  if (!target) return undefined;
  const fileName = target.split('/').pop() ?? target;
  return (
    mediaByPath[target] ??
    mediaByPath[`${rootDir}/${target.replace(new RegExp(`^${rootDir}/`), '')}`] ??
    mediaByName[fileName]
  );
}

export function collectMedia(entries: OfficeEntryMap, mediaPrefix: string) {
  const media = createMediaStore();

  for (const [path] of entries) {
    if (!path.startsWith(mediaPrefix)) continue;
    const binary = readBinary(entries, path);
    if (!binary) continue;
    media.register(path, binary);
  }

  return media.store;
}
