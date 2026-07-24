import {
  DEFAULT_DOC_BLOCK_TARGET_BYTES,
  estimateDocBlockBytes,
} from './chunkDocBlocks';
import type { DocBlock } from './types';

export type DocBlockBatch = {
  startIndex: number;
  blocks: DocBlock[];
};

export type DocBlockStreamOptions = {
  targetBytes?: number;
  onBatch?(batch: DocBlockBatch): Promise<void>;
};

function blockHasImage(block: DocBlock) {
  if (block.type === 'paragraph') {
    return Boolean(block.inlines?.some((inline) => inline.type === 'image'));
  }
  if (block.type === 'table') {
    return block.rows.some((row) =>
      row.cells.some((cell) =>
        cell.inlines?.some((inline) => inline.type === 'image'),
      ),
    );
  }
  return block.items.some((item) =>
    item.inlines?.some((inline) => inline.type === 'image'),
  );
}

function isImageOnlyParagraph(block: DocBlock) {
  return (
    block.type === 'paragraph' && !block.text.trim() && blockHasImage(block)
  );
}

function isShapeTextParagraph(block: DocBlock) {
  if (block.type !== 'paragraph') return false;
  const text = block.text.replace(/\s+/g, '');
  return (
    !blockHasImage(block) &&
    Boolean(text) &&
    (text.includes('\u6dfb\u52a0\u6807\u9898') ||
      text.includes('\u8bf7\u70b9\u51fb\u7f16\u8f91\u6587\u5b57') ||
      text.includes('\u8bf7\u6b64\u5904\u7f16\u8f91\u6587\u5b57'))
  );
}

function finalizeBlockId(block: DocBlock, index: number): DocBlock {
  const prefix =
    block.type === 'table'
      ? 'doc-table'
      : block.type === 'list'
      ? 'doc-list'
      : 'doc-p';
  return { ...block, id: `${prefix}-${index + 1}` };
}

/** 以单遍方式保持 DOC 局部重排语义并产生稳定传输批次。 */
export class DocBlockStreamBuilder {
  private readonly targetBytes: number;
  private readonly onBatch?: DocBlockStreamOptions['onBatch'];
  private readonly finalBlocks: DocBlock[] = [];
  private batchStartIndex = 0;
  private batchBlocks: DocBlock[] = [];
  private batchBytes = 0;
  private sourceCount = 0;
  private pendingTable: DocBlock | undefined;
  private pendingImages: DocBlock[] = [];
  private pendingShapeTexts: DocBlock[] = [];

  constructor(options: DocBlockStreamOptions = {}) {
    this.targetBytes = Math.max(
      1,
      options.targetBytes ?? DEFAULT_DOC_BLOCK_TARGET_BYTES,
    );
    this.onBatch = options.onBatch;
  }

  get nextSourceIndex() {
    return this.sourceCount;
  }

  async add(block: DocBlock): Promise<void> {
    this.sourceCount += 1;
    await this.accept(block);
  }

  async finish(): Promise<DocBlock[]> {
    await this.flushPendingGroup();
    await this.flushBatch();
    return [...this.finalBlocks];
  }

  private async accept(block: DocBlock): Promise<void> {
    if (!this.pendingTable) {
      if (block.type === 'table') {
        this.pendingTable = block;
        return;
      }
      await this.finalize(block);
      return;
    }
    if (!this.pendingShapeTexts.length && isImageOnlyParagraph(block)) {
      this.pendingImages.push(block);
      return;
    }
    if (this.pendingImages.length && isShapeTextParagraph(block)) {
      this.pendingShapeTexts.push(block);
      return;
    }
    await this.flushPendingGroup();
    await this.accept(block);
  }

  private async flushPendingGroup(): Promise<void> {
    if (!this.pendingTable) return;
    const reordered =
      this.pendingImages.length && this.pendingShapeTexts.length
        ? [this.pendingTable, ...this.pendingShapeTexts, ...this.pendingImages]
        : [this.pendingTable, ...this.pendingImages, ...this.pendingShapeTexts];
    this.pendingTable = undefined;
    this.pendingImages = [];
    this.pendingShapeTexts = [];
    for (const block of reordered) {
      await this.finalize(block);
    }
  }

  private async finalize(source: DocBlock): Promise<void> {
    const block = finalizeBlockId(source, this.finalBlocks.length);
    const blockBytes = estimateDocBlockBytes(block);
    if (
      this.batchBlocks.length &&
      this.batchBytes + blockBytes > this.targetBytes
    ) {
      await this.flushBatch();
    }
    if (!this.batchBlocks.length) {
      this.batchStartIndex = this.finalBlocks.length;
    }
    this.finalBlocks.push(block);
    this.batchBlocks.push(block);
    this.batchBytes += blockBytes;
    if (this.batchBytes >= this.targetBytes) {
      await this.flushBatch();
    }
  }

  private async flushBatch(): Promise<void> {
    if (!this.batchBlocks.length) return;
    const batch = {
      startIndex: this.batchStartIndex,
      blocks: this.batchBlocks,
    };
    this.batchBlocks = [];
    this.batchBytes = 0;
    await this.onBatch?.(batch);
  }
}
