import type { PortableResource } from '../protocol/messages';
import { readResourceReference } from './resourceReferences';

function rgbaToPngBlob(resource: Extract<PortableResource, { encoding: 'rgba' }>) {
  if (typeof document === 'undefined') {
    throw new Error('当前环境没有 Canvas，无法转换 DIB');
  }
  const canvas = document.createElement('canvas');
  canvas.width = resource.width;
  canvas.height = resource.height;
  const context = canvas.getContext('2d');
  if (!context) throw new Error('无法创建 DIB Canvas 上下文');
  const pixels = new Uint8ClampedArray(resource.buffer);
  const imageData = context.createImageData(resource.width, resource.height);
  imageData.data.set(pixels);
  context.putImageData(imageData, 0, 0);
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error('DIB 转 PNG 失败'));
    }, 'image/png');
  });
}

async function resourceToBlob(resource: PortableResource) {
  if (resource.encoding === 'binary') {
    return new Blob([resource.buffer], { type: resource.mimeType });
  }
  if (resource.encoding === 'text') {
    return new Blob([resource.text], {
      type: 'image/svg+xml;charset=utf-8',
    });
  }
  return rgbaToPngBlob(resource);
}

/** 在主线程创建和管理解析资源的 Blob URL。 */
export class ResourceRegistry {
  private readonly urls = new Map<string, string>();
  private readonly ownedUrls = new Set<string>();

  async register(resource: PortableResource): Promise<string> {
    const existing = this.urls.get(resource.id);
    if (existing) return existing;
    if (
      typeof URL === 'undefined' ||
      typeof URL.createObjectURL !== 'function'
    ) {
      throw new Error('当前环境不支持 Blob URL');
    }
    const url = URL.createObjectURL(await resourceToBlob(resource));
    this.urls.set(resource.id, url);
    this.ownedUrls.add(url);
    return url;
  }

  resolve(reference: string): string {
    const resourceId = readResourceReference(reference);
    if (!resourceId) return reference;
    const url = this.urls.get(resourceId);
    if (!url) {
      const error = new Error(`解析资源不存在：${resourceId}`) as Error & {
        code: string;
      };
      error.code = 'RESOURCE_NOT_FOUND';
      throw error;
    }
    return url;
  }

  /** 将 URL 的释放责任移交给最终文档。 */
  takeObjectUrls(): string[] {
    const urls = [...this.ownedUrls];
    this.ownedUrls.clear();
    return urls;
  }

  dispose() {
    if (typeof URL !== 'undefined' && typeof URL.revokeObjectURL === 'function') {
      this.ownedUrls.forEach((url) => URL.revokeObjectURL(url));
    }
    this.ownedUrls.clear();
    this.urls.clear();
  }
}
