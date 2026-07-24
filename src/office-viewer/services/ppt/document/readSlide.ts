import type { ThemeModel } from '../../presentation/types';
import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { parsePptDrawing } from '../drawing';
import type {
  PptEditChain,
  PptParseContext,
  PptSlideDescriptor,
  PptSlideModel,
} from '../types';

/** 读取一页幻灯片的母版引用、文本与 OfficeArt 绘图。 */
export function readPptSlide(
  documentStream: Uint8Array,
  editChain: PptEditChain,
  descriptor: PptSlideDescriptor,
  width: number,
  height: number,
  theme: ThemeModel,
  context: PptParseContext,
): PptSlideModel | undefined {
  const offset = editChain.persistOffsets.get(descriptor.persistId);
  if (offset === undefined) {
    context.warnings.push({
      code: 'PPT_SLIDE_MISSING',
      message: `持久化目录中缺少第 ${descriptor.index} 页`,
      slideIndex: descriptor.index,
    });
    return undefined;
  }

  const record = new PptRecordReader(
    documentStream,
    offset,
    documentStream.length,
  ).readRecord();
  if (!record || record.type !== PPT_RECORD.SLIDE) {
    context.warnings.push({
      code: 'PPT_SLIDE_CORRUPT',
      message: `第 ${descriptor.index} 页不是有效的 SlideContainer`,
      slideIndex: descriptor.index,
      offset,
    });
    return undefined;
  }

  let masterId: number | undefined;
  let drawing: Uint8Array | undefined;
  const children = new PptRecordReader(
    documentStream,
    record.dataOffset,
    record.endOffset,
  );
  for (const child of children.records()) {
    if (child.type === PPT_RECORD.SLIDE_ATOM && child.length >= 16) {
      const view = new DataView(
        child.data.buffer,
        child.data.byteOffset,
        child.data.byteLength,
      );
      masterId = view.getUint32(12, true);
    }
    if (child.type === PPT_RECORD.PP_DRAWING) drawing = child.data;
  }

  return {
    id: `ppt-slide-${descriptor.persistId}`,
    persistId: descriptor.persistId,
    slideId: descriptor.slideId,
    index: descriptor.index,
    width,
    height,
    masterId,
    hidden: descriptor.hidden,
    elements: drawing ? parsePptDrawing(drawing, theme, context) : [],
    sourceOffset: offset,
  };
}
