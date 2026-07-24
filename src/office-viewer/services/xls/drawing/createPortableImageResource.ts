import type { PortableResource } from '../../parsing/protocol/messages';
import { createResourceReference } from '../../parsing/assembly/resourceReferences';
import type { SpreadsheetWarning } from '../../spreadsheet/types';
import { decodeDib } from './decodeDib';
import { parseEmf } from './metafile/parseEmf';
import { parseWmf } from './metafile/parseWmf';
import { vectorSceneToSvg } from './metafile/vectorSceneToSvg';
import type { Biff8DrawingImage } from './types';

export type PortableImageResourceResult = {
  reference: string;
  resource: PortableResource;
  warnings: SpreadsheetWarning[];
};

function toExactArrayBuffer(bytes: ArrayLike<number>) {
  return Uint8Array.from(bytes).buffer;
}

async function inflateMetafile(
  image: Biff8DrawingImage,
  warnings: SpreadsheetWarning[],
) {
  if (!image.compressed) return image.bytes;
  if (typeof DecompressionStream === 'undefined') {
    warnings.push({
      code: 'DEFLATE_UNAVAILABLE',
      message: '当前浏览器不支持 DecompressionStream，无法显示压缩 metafile',
    });
    throw new Error('浏览器缺少 DecompressionStream');
  }
  const stream = new Blob([image.bytes])
    .stream()
    .pipeThrough(new DecompressionStream('deflate'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function rasterMimeType(format: Biff8DrawingImage['format']) {
  if (format === 'png') return 'image/png';
  if (format === 'jpeg') return 'image/jpeg';
  if (format === 'gif') return 'image/gif';
  return undefined;
}

/** 将 BIFF8 图片转换成可跨线程传输的资源，不创建浏览器 URL。 */
export async function createPortableImageResource(
  image: Biff8DrawingImage,
  resourceId = image.id,
): Promise<PortableImageResourceResult> {
  const warnings = [...image.warnings];
  const mimeType = rasterMimeType(image.format);
  let resource: PortableResource;
  if (mimeType) {
    resource = {
      id: resourceId,
      encoding: 'binary',
      mimeType,
      buffer: toExactArrayBuffer(image.bytes),
    };
  } else if (image.format === 'dib') {
    const bitmap = decodeDib(image.bytes);
    resource = {
      id: resourceId,
      encoding: 'rgba',
      mimeType: 'image/png',
      width: bitmap.width,
      height: bitmap.height,
      buffer: toExactArrayBuffer(bitmap.rgba),
    };
  } else if (image.format === 'wmf' || image.format === 'emf') {
    const bytes = await inflateMetafile(image, warnings);
    const scene = image.format === 'wmf' ? parseWmf(bytes) : parseEmf(bytes);
    warnings.push(...scene.warnings);
    resource = {
      id: resourceId,
      encoding: 'text',
      mimeType: 'image/svg+xml',
      text: vectorSceneToSvg(scene),
    };
  } else {
    throw new Error(`暂不支持 ${image.format} 图片`);
  }
  return {
    reference: createResourceReference(resourceId),
    resource,
    warnings,
  };
}
