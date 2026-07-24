import { ResourceRegistry } from '../../parsing/assembly/ResourceRegistry';
import { createPortableImageResource } from './createPortableImageResource';
import type { Biff8DrawingImage } from './types';

/** 将绘图图片转换为浏览器可显示的 Blob URL。 */
export async function createImageResource(image: Biff8DrawingImage) {
  const portable = await createPortableImageResource(image);
  const registry = new ResourceRegistry();
  const objectUrl = await registry.register(portable.resource);
  registry.takeObjectUrls();
  return {
    src: objectUrl,
    objectUrl,
    warnings: portable.warnings,
  };
}
