import type { PresentationDocument } from './types';

/** 释放演示文稿解析期间创建的 Blob URL；重复调用保持幂等。 */
export function disposePresentationDocument(
  document: PresentationDocument | undefined,
) {
  const urls = document?.resources?.objectUrls;
  if (!urls?.length) return;
  const uniqueUrls = new Set(urls);
  urls.length = 0;
  if (typeof URL === 'undefined' || typeof URL.revokeObjectURL !== 'function') {
    return;
  }
  uniqueUrls.forEach((url) => URL.revokeObjectURL(url));
}
