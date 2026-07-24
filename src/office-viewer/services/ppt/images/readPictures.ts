import {
  OFFICE_ART_RECORD,
  parseOfficeArtRecords,
  type OfficeArtRecord,
} from '../../../shared/officeart';
import {
  createPptResourceId,
  registerPptResource,
  type PptParseContext,
} from '../types';
import { createPptStaticPreviewCard } from './createStaticPreviewCard';

type RasterInfo = {
  mimeType: string;
  bytes: Uint8Array;
};

function findSignature(bytes: Uint8Array, signature: number[]) {
  const limit = Math.min(bytes.length - signature.length, 96);
  for (let offset = 0; offset <= limit; offset += 1) {
    if (signature.every((value, index) => bytes[offset + index] === value)) {
      return offset;
    }
  }
  return -1;
}

function readRaster(record: OfficeArtRecord): RasterInfo | undefined {
  const candidates = [
    {
      type: OFFICE_ART_RECORD.BLIP_PNG,
      mimeType: 'image/png',
      signature: [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a],
    },
    {
      type: OFFICE_ART_RECORD.BLIP_JPEG,
      mimeType: 'image/jpeg',
      signature: [0xff, 0xd8, 0xff],
    },
  ];
  const candidate = candidates.find((item) => item.type === record.type);
  if (!candidate) return undefined;
  const offset = findSignature(record.data, candidate.signature);
  if (offset < 0) return undefined;
  return {
    mimeType: candidate.mimeType,
    bytes: record.data.subarray(offset),
  };
}

function fallbackLabel(type: number) {
  if (type === OFFICE_ART_RECORD.BLIP_WMF) return ['WMF 图像', '矢量图静态预览'];
  if (type === OFFICE_ART_RECORD.BLIP_EMF) return ['EMF 图像', '矢量图静态预览'];
  if (type === OFFICE_ART_RECORD.BLIP_PICT) return ['PICT 图像', '嵌入图像静态预览'];
  if (type === OFFICE_ART_RECORD.BLIP_DIB) return ['DIB 图像', '位图静态预览'];
  return ['嵌入图像', 'PowerPoint 97–2003 图片对象'];
}

function toExactArrayBuffer(bytes: Uint8Array) {
  return Uint8Array.from(bytes).buffer;
}

/** 解析 Pictures 流并建立一基序号到可传输资源引用的映射。 */
export async function readPptPictures(
  picturesStream: Uint8Array | undefined,
  context: PptParseContext,
) {
  if (!picturesStream?.length) return context.blipUrls;
  const records = parseOfficeArtRecords(picturesStream, context.warnings);
  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];
    const raster = readRaster(record);
    let reference: string;
    if (raster) {
      reference = registerPptResource(context, {
        id: createPptResourceId(context, 'picture'),
        encoding: 'binary',
        mimeType: raster.mimeType,
        buffer: toExactArrayBuffer(raster.bytes),
      });
    } else {
      const [title, detail] = fallbackLabel(record.type);
      reference = createPptStaticPreviewCard(title, detail, context);
    }
    context.blipUrls.set(index + 1, reference);
    await context.yieldIfNeeded();
  }
  return context.blipUrls;
}
