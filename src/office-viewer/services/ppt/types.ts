import type {
  PresentationWarning,
  SlideBackground,
  SlideElement,
  TextStyle,
  ThemeModel,
} from '../presentation/types';
import { createResourceReference } from '../parsing/assembly/resourceReferences';
import type { PortableResource } from '../parsing/protocol/messages';
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
  resources: PortableResource[];
  resourceSequence: number;
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

/** 创建一次解析独占的 warning、可传输资源与时间片上下文。 */
export function createPptParseContext(
  yieldIfNeeded: () => Promise<void>,
): PptParseContext {
  return {
    warnings: [],
    resources: [],
    resourceSequence: 0,
    blipUrls: new Map<number, string>(),
    charts: new Map<number, { chart: OfficeChartModel; title?: string }>(),
    yieldIfNeeded,
  };
}

/** 为 PPT 资源分配会话内稳定且不冲突的标识。 */
export function createPptResourceId(
  context: PptParseContext,
  category: string,
) {
  context.resourceSequence += 1;
  return `ppt:${category}:${context.resourceSequence}`;
}

/** 注册可传输资源，并返回只在解析模型中使用的资源引用。 */
export function registerPptResource(
  context: PptParseContext,
  resource: PortableResource,
) {
  context.resources.push(resource);
  return createResourceReference(resource.id);
}
