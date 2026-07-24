import type { SpreadsheetWarning } from '../../spreadsheet/types';
import { decodeDib } from './decodeDib';
import { parseEmf } from './metafile/parseEmf';
import { parseWmf } from './metafile/parseWmf';
import { vectorSceneToSvg } from './metafile/vectorSceneToSvg';
import type { Biff8DrawingImage } from './types';

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

async function dibToPngBlob(image: Biff8DrawingImage) {
  if (typeof document === 'undefined') {
    throw new Error('当前环境没有 Canvas，无法转换 DIB');
  }
  const decoded = decodeDib(image.bytes);
  const canvas = document.createElement('canvas');
  canvas.width = decoded.width;
  canvas.height = decoded.height;
  const context = canvas.getContext('2d');
  if (!context) throw new Error('无法创建 DIB Canvas 上下文');
  const imageData = context.createImageData(decoded.width, decoded.height);
  imageData.data.set(decoded.rgba);
  context.putImageData(imageData, 0, 0);
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error('DIB 转 PNG 失败'));
    }, 'image/png');
  });
}

function createObjectUrl(blob: Blob) {
  if (typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function') {
    throw new Error('当前环境不支持 Blob URL');
  }
  return URL.createObjectURL(blob);
}

/** 将绘图图片转换为浏览器可显示的 Blob URL。 */
export async function createImageResource(image: Biff8DrawingImage) {
  const warnings = [...image.warnings];
  let blob: Blob;
  if (
    image.format === 'png' ||
    image.format === 'jpeg' ||
    image.format === 'gif'
  ) {
    const mimeType =
      image.format === 'png'
        ? 'image/png'
        : image.format === 'jpeg'
        ? 'image/jpeg'
        : 'image/gif';
    blob = new Blob([image.bytes], { type: mimeType });
  } else if (image.format === 'dib') {
    blob = await dibToPngBlob(image);
  } else if (image.format === 'wmf' || image.format === 'emf') {
    const bytes = await inflateMetafile(image, warnings);
    const scene = image.format === 'wmf' ? parseWmf(bytes) : parseEmf(bytes);
    warnings.push(...scene.warnings);
    blob = new Blob([vectorSceneToSvg(scene)], {
      type: 'image/svg+xml;charset=utf-8',
    });
  } else {
    throw new Error(`暂不支持 ${image.format} 图片`);
  }
  const objectUrl = createObjectUrl(blob);
  return { src: objectUrl, objectUrl, warnings };
}
