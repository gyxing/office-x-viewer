import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import type {
  PptParseContext,
  PptSlideDescriptor,
  PptRecord,
} from '../types';

export type PptMasterDescriptor = {
  persistId: number;
  masterId: number;
};

function readSlidePersistFields(record: PptRecord) {
  if (record.length < 16) return undefined;
  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  return {
    persistId: view.getUint32(0, true),
    objectId: view.getUint32(12, true),
  };
}

/** 从指定实例的 SlideListWithTextContainer 恢复页面或母版引用顺序。 */
export function readPptSlideLists(
  documentStream: Uint8Array,
  documentRecord: PptRecord,
  context: PptParseContext,
) {
  const slides: PptSlideDescriptor[] = [];
  const masters: PptMasterDescriptor[] = [];
  const reader = new PptRecordReader(
    documentStream,
    documentRecord.dataOffset,
    documentRecord.endOffset,
  );

  for (const child of reader.records()) {
    if (child.type !== PPT_RECORD.SLIDE_LIST_WITH_TEXT) continue;
    if (child.instance !== 0 && child.instance !== 1) continue;
    const listReader = new PptRecordReader(
      documentStream,
      child.dataOffset,
      child.endOffset,
    );
    for (const item of listReader.records()) {
      if (item.type !== PPT_RECORD.SLIDE_PERSIST_ATOM) continue;
      const fields = readSlidePersistFields(item);
      if (!fields?.persistId) {
        context.warnings.push({
          code: 'PPT_SLIDE_PERSIST_INVALID',
          message: '幻灯片或母版引用记录长度无效',
          offset: item.offset,
        });
        continue;
      }
      if (child.instance === 0) {
        slides.push({
          persistId: fields.persistId,
          slideId: fields.objectId,
          index: slides.length + 1,
        });
      } else {
        masters.push({
          persistId: fields.persistId,
          masterId: fields.objectId,
        });
      }
    }
  }

  return { slides, masters };
}
