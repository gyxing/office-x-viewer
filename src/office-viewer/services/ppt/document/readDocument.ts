import type { ThemeModel } from '../../presentation/types';
import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { PptParseError } from '../errors';
import type {
  PptBinaryDocument,
  PptEditChain,
  PptParseContext,
  PptSlideModel,
} from '../types';
import { readPptMaster } from './readMaster';
import { readPptSlide } from './readSlide';
import { readPptSlideLists } from './readSlideLists';

const DEFAULT_SLIDE_WIDTH = 960;
const DEFAULT_SLIDE_HEIGHT = 540;
const MASTER_UNIT_TO_PX = 1 / 8;

export type PptDocumentStructure = Pick<
  PptBinaryDocument,
  'width' | 'height' | 'theme' | 'masters'
>;

export type PptDocumentObserver = {
  structure(value: PptDocumentStructure): Promise<void>;
  slide(index: number, slide: PptSlideModel): Promise<void>;
};

function readDocumentSize(
  documentStream: Uint8Array,
  documentRecordOffset: number,
) {
  const documentRecord = new PptRecordReader(
    documentStream,
    documentRecordOffset,
    documentStream.length,
  ).readRecord()!;
  const reader = new PptRecordReader(
    documentStream,
    documentRecord.dataOffset,
    documentRecord.endOffset,
  );
  for (const child of reader.records()) {
    if (child.type !== PPT_RECORD.DOCUMENT_ATOM || child.length < 8) continue;
    const view = new DataView(
      child.data.buffer,
      child.data.byteOffset,
      child.data.byteLength,
    );
    const width = view.getInt32(0, true) * MASTER_UNIT_TO_PX;
    const height = view.getInt32(4, true) * MASTER_UNIT_TO_PX;
    if (width > 0 && height > 0) return { width, height };
  }
  return { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT };
}

/** 从最终 persist 目录恢复文档、母版和正式幻灯片顺序。 */
export async function readPptBinaryDocument(
  documentStream: Uint8Array,
  editChain: PptEditChain,
  context: PptParseContext,
  observer?: PptDocumentObserver,
): Promise<PptBinaryDocument> {
  const documentOffset = editChain.persistOffsets.get(
    editChain.documentPersistId,
  )!;
  const documentRecord = new PptRecordReader(
    documentStream,
    documentOffset,
    documentStream.length,
  ).readRecord();
  if (!documentRecord || documentRecord.type !== PPT_RECORD.DOCUMENT) {
    throw new PptParseError(
      'PPT_DOCUMENT_MISSING',
      '无法读取 PowerPoint 根文档对象',
      { offset: documentOffset, recordType: documentRecord?.type },
    );
  }

  const { width, height } = readDocumentSize(documentStream, documentOffset);
  const theme: ThemeModel = {
    colorScheme: {
      lt1: '#ffffff',
      dk1: '#000000',
      accent1: '#4472c4',
      accent2: '#ed7d31',
    },
    fontScheme: {},
    colorMap: {
      bg1: 'lt1',
      tx1: 'dk1',
      accent1: 'accent1',
      accent2: 'accent2',
    },
  };
  const descriptors = readPptSlideLists(
    documentStream,
    documentRecord,
    context,
  );
  const masters = new Map(
    descriptors.masters
      .map((descriptor) => {
        const master = readPptMaster(
          documentStream,
          editChain,
          descriptor,
          theme,
          context,
        );
        return master ? ([master.id, master] as const) : undefined;
      })
      .filter(
        (entry): entry is readonly [number, NonNullable<typeof entry>[1]] =>
          Boolean(entry),
      ),
  );
  await observer?.structure({ width, height, theme, masters });
  const slides = [];
  for (const descriptor of descriptors.slides) {
    const slide = readPptSlide(
      documentStream,
      editChain,
      descriptor,
      width,
      height,
      theme,
      context,
    );
    if (slide) {
      slide.background =
        masters.get(slide.masterId ?? Number.NaN)?.background ?? {
          fill: theme.colorScheme.lt1 ?? '#ffffff',
      };
      slides.push(slide);
      await observer?.slide(slides.length - 1, slide);
    }
    await context.yieldIfNeeded();
  }
  if (!slides.length) {
    throw new PptParseError(
      'PPT_NO_VALID_SLIDES',
      'PPT 文件中没有可预览的有效幻灯片',
    );
  }

  return {
    width,
    height,
    theme,
    masters,
    slides,
    externalObjects: new Map(),
    warnings: context.warnings,
  };
}
