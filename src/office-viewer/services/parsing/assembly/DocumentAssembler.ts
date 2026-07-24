import type {
  SpreadsheetSheet,
  SpreadsheetWarning,
  SpreadsheetWorkbook,
} from '../../spreadsheet/types';
import type {
  PresentationDocument,
  SlideElement,
  SlideModel,
} from '../../presentation/types';
import type {
  PortablePresentationMetadata,
  PortableDocMetadata,
  PortableResource,
} from '../protocol/messages';
import type {
  DocBlock,
  DocDocument,
  DocImage,
  DocTextInline,
} from '../../doc/types';
import { paragraphsFromDocBlocks } from '../../doc/chunkDocBlocks';
import { ResourceRegistry } from './ResourceRegistry';

type VersionedSpreadsheetSheet = {
  revision: number;
  sheet: SpreadsheetSheet;
};

/** 将 XLS 跨线程分块还原成现有电子表格模型。 */
export class XlsDocumentAssembler {
  private readonly sheets = new Map<number, VersionedSpreadsheetSheet>();
  private warnings: SpreadsheetWarning[] = [];
  private completed = false;

  constructor(private readonly resources: ResourceRegistry) {}

  async addResource(resource: PortableResource) {
    await this.resources.register(resource);
  }

  addSheet(index: number, revision: number, sheet: SpreadsheetSheet) {
    if (this.completed) throw new Error('XLS 组装已经完成');
    const current = this.sheets.get(index);
    if (current && revision <= current.revision) {
      throw new Error(`XLS 工作表修订无效：${index}@${revision}`);
    }
    sheet.images.forEach((image) => {
      image.src = this.resources.resolve(image.src);
    });
    this.sheets.set(index, { revision, sheet });
  }

  setWarnings(warnings: SpreadsheetWarning[] | undefined) {
    this.warnings = warnings ? [...warnings] : [];
  }

  hasRenderableContent() {
    return this.sheets.size > 0;
  }

  snapshot(): SpreadsheetWorkbook {
    if (!this.hasRenderableContent()) {
      throw new Error('XLS 组装尚无可渲染工作表');
    }
    return this.createWorkbook();
  }

  complete(): SpreadsheetWorkbook {
    return this.finish();
  }

  completePartial(): SpreadsheetWorkbook {
    if (!this.hasRenderableContent()) {
      throw new Error('XLS 组装尚无可保留内容');
    }
    return this.finish();
  }

  private createWorkbook(objectUrls: string[] = []): SpreadsheetWorkbook {
    return {
      sheets: [...this.sheets.entries()]
        .sort(([left], [right]) => left - right)
        .map(([, value]) => value.sheet),
      warnings: this.warnings.length ? [...this.warnings] : undefined,
      resources: objectUrls.length ? { objectUrls } : undefined,
    };
  }

  private finish() {
    if (this.completed) throw new Error('XLS 组装已经完成');
    this.completed = true;
    return this.createWorkbook(this.resources.takeObjectUrls());
  }

  dispose() {
    if (!this.completed) this.resources.dispose();
    this.sheets.clear();
    this.warnings = [];
  }
}

function resolveElementResources(
  element: SlideElement,
  resources: ResourceRegistry,
) {
  if (element.type === 'image') {
    element.src = resources.resolve(element.src);
    return;
  }
  if (element.type === 'group') {
    element.children.forEach((child) =>
      resolveElementResources(child, resources),
    );
  }
}

function resolveSlideResources(
  slide: SlideModel,
  resources: ResourceRegistry,
) {
  if (slide.background?.imageRef) {
    slide.background.imageRef = resources.resolve(
      slide.background.imageRef,
    );
  }
  slide.elements.forEach((element) =>
    resolveElementResources(element, resources),
  );
}

/** 将 PPT 跨线程元数据、资源和幻灯片还原成现有演示文稿模型。 */
export class PptDocumentAssembler {
  private metadata: PortablePresentationMetadata | undefined;
  private readonly slides = new Map<number, SlideModel>();
  private completed = false;

  constructor(private readonly resources: ResourceRegistry) {}

  async addResource(resource: PortableResource) {
    await this.resources.register(resource);
  }

  setMetadata(metadata: PortablePresentationMetadata) {
    if (this.completed) throw new Error('PPT 组装已经完成');
    this.metadata = metadata;
  }

  addSlide(index: number, slide: SlideModel) {
    if (this.completed) throw new Error('PPT 组装已经完成');
    resolveSlideResources(slide, this.resources);
    this.slides.set(index, slide);
  }

  hasRenderableContent() {
    return Boolean(this.metadata && this.slides.size);
  }

  snapshot(): PresentationDocument {
    if (!this.hasRenderableContent()) {
      throw new Error('PPT 组装尚无可渲染幻灯片');
    }
    return this.createDocument();
  }

  complete(): PresentationDocument {
    return this.finish();
  }

  completePartial(): PresentationDocument {
    if (!this.hasRenderableContent()) {
      throw new Error('PPT 组装尚无可保留内容');
    }
    return this.finish();
  }

  private createDocument(objectUrls: string[] = []): PresentationDocument {
    if (!this.metadata) throw new Error('PPT 组装缺少演示文稿元数据');
    return {
      ...this.metadata,
      slides: [...this.slides.entries()]
        .sort(([left], [right]) => left - right)
        .map(([, slide]) => slide),
      resources: objectUrls.length ? { objectUrls } : undefined,
    };
  }

  private finish(): PresentationDocument {
    if (this.completed) throw new Error('PPT 组装已经完成');
    if (!this.metadata) throw new Error('PPT 组装缺少演示文稿元数据');
    const document = this.createDocument(this.resources.takeObjectUrls());
    this.completed = true;
    return document;
  }

  dispose() {
    if (!this.completed) this.resources.dispose();
    this.metadata = undefined;
    this.slides.clear();
  }
}

function resolveDocImage(image: DocImage, resources: ResourceRegistry) {
  image.src = resources.resolve(image.src);
}

function resolveDocInlines(
  inlines: DocTextInline[] | undefined,
  resources: ResourceRegistry,
) {
  inlines?.forEach((inline) => {
    if (inline.type === 'image') resolveDocImage(inline.image, resources);
  });
}

function resolveDocBlockResources(
  block: DocBlock,
  resources: ResourceRegistry,
) {
  if (block.type === 'paragraph') {
    resolveDocInlines(block.inlines, resources);
    return;
  }
  if (block.type === 'list') {
    block.items.forEach((item) => resolveDocInlines(item.inlines, resources));
    return;
  }
  block.rows.forEach((row) => {
    row.cells.forEach((cell) => resolveDocInlines(cell.inlines, resources));
  });
}

/** 将 DOC 跨线程元数据、正文块和图片资源还原成现有文档模型。 */
export class DocDocumentAssembler {
  private metadata: PortableDocMetadata | undefined;
  private readonly blocks = new Map<number, DocBlock>();
  private completed = false;

  constructor(private readonly resources: ResourceRegistry) {}

  async addResource(resource: PortableResource) {
    await this.resources.register(resource);
  }

  setMetadata(metadata: PortableDocMetadata) {
    if (this.completed) throw new Error('DOC 组装已经完成');
    const images = metadata.images.map((image) => ({ ...image }));
    images.forEach((image) => resolveDocImage(image, this.resources));
    this.metadata = { ...metadata, images };
  }

  addBlocks(startIndex: number, blocks: DocBlock[]) {
    if (this.completed) throw new Error('DOC 组装已经完成');
    blocks.forEach((block, offset) => {
      const index = startIndex + offset;
      if (this.blocks.has(index)) {
        throw new Error(`DOC 正文块索引重复：${index}`);
      }
      resolveDocBlockResources(block, this.resources);
      this.blocks.set(index, block);
    });
  }

  hasRenderableContent() {
    return Boolean(this.metadata && this.blocks.size);
  }

  snapshot(): DocDocument {
    if (!this.hasRenderableContent()) {
      throw new Error('DOC 组装尚无可渲染正文');
    }
    return this.createDocument();
  }

  complete(): DocDocument {
    return this.finish();
  }

  completePartial(): DocDocument {
    if (!this.hasRenderableContent()) {
      throw new Error('DOC 组装尚无可保留内容');
    }
    return this.finish();
  }

  private createDocument(objectUrls: string[] = []): DocDocument {
    if (!this.metadata) throw new Error('DOC 组装缺少文档元数据');
    const blocks = [...this.blocks.entries()]
      .sort(([left], [right]) => left - right)
      .map(([, block]) => block);
    return {
      ...this.metadata,
      blocks,
      paragraphs: paragraphsFromDocBlocks(blocks),
      resources: objectUrls.length ? { objectUrls } : undefined,
    };
  }

  private finish(): DocDocument {
    if (this.completed) throw new Error('DOC 组装已经完成');
    if (!this.metadata) throw new Error('DOC 组装缺少文档元数据');
    const document = this.createDocument(this.resources.takeObjectUrls());
    this.completed = true;
    return document;
  }

  dispose() {
    if (!this.completed) this.resources.dispose();
    this.metadata = undefined;
    this.blocks.clear();
  }
}
