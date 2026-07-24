import type {
  PresentationWarning,
  SlideBackground,
  SlideElement,
  TextStyle,
  ThemeModel,
} from '../presentation/types';
import type { OfficeChartModel } from '../../shared/ooxml/charts';

export type PptRecord = {
  version: number;
  instance: number;
  type: number;
  length: number;
  offset: number;
  dataOffset: number;
  endOffset: number;
  data: Uint8Array;
};

export type PptParseContext = {
  warnings: PresentationWarning[];
  objectUrls: Set<string>;
  blipUrls: Map<number, string>;
  charts: Map<number, { chart: OfficeChartModel; title?: string }>;
  yieldIfNeeded: () => Promise<void>;
};

export type PptPersistObjectMap = Map<number, number>;

export type PptEditChain = {
  documentPersistId: number;
  persistIdSeed: number;
  persistOffsets: PptPersistObjectMap;
  editOffsets: number[];
};

export type PptSlideDescriptor = {
  persistId: number;
  slideId: number;
  index: number;
  hidden?: boolean;
};

export type PptMasterModel = {
  id: number;
  persistId: number;
  background?: SlideBackground;
  textDefaults?: TextStyle;
  elements: SlideElement[];
};

export type PptExternalObject = {
  id: number;
  persistId?: number;
  name?: string;
  type?: string;
  previewBlipId?: number;
};

export type PptSlideModel = {
  id: string;
  persistId: number;
  slideId: number;
  index: number;
  width: number;
  height: number;
  masterId?: number;
  hidden?: boolean;
  background?: SlideBackground;
  elements: SlideElement[];
  sourceOffset: number;
};

export type PptBinaryDocument = {
  width: number;
  height: number;
  theme: ThemeModel;
  masters: Map<number, PptMasterModel>;
  slides: PptSlideModel[];
  externalObjects: Map<number, PptExternalObject>;
  warnings: PresentationWarning[];
};

/** 创建一次解析独占的 warning、Blob URL 与时间片上下文。 */
export function createPptParseContext(
  yieldIfNeeded: () => Promise<void>,
  objectUrls = new Set<string>(),
): PptParseContext {
  return {
    warnings: [],
    objectUrls,
    blipUrls: new Map<number, string>(),
    charts: new Map<number, { chart: OfficeChartModel; title?: string }>(),
    yieldIfNeeded,
  };
}
