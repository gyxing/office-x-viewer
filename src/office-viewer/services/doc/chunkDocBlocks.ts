import type { PortableDocMetadata } from '../parsing/protocol/messages';
import type { DocBlock, DocDocument, DocParagraph } from './types';

export type DocBlockChunk = {
  startIndex: number;
  blocks: DocBlock[];
};

/** DOC 正文批次默认按 256 KiB 估算，单个文档块始终保持完整。 */
export const DEFAULT_DOC_BLOCK_TARGET_BYTES = 256 * 1024;

/** 从最终 DOC 模型中提取适合先行传输的页面、摘要和图片元数据。 */
export function documentMetadataFromDoc(
  document: DocDocument,
): PortableDocMetadata {
  const { blocks, paragraphs, resources, ...metadata } = document;
  void blocks;
  void paragraphs;
  void resources;
  return metadata;
}

/** 从正文块重建兼容旧渲染器使用的扁平段落摘要。 */
export function paragraphsFromDocBlocks(blocks: DocBlock[]): DocParagraph[] {
  return blocks
    .flatMap((block) => {
      if (block.type === 'paragraph') return [block.text];
      if (block.type === 'list') return block.items.map((item) => item.text);
      return block.rows.map((row) =>
        row.cells.map((cell) => cell.text).join(' '),
      );
    })
    .filter(Boolean)
    .map((text, index) => ({
      id: `doc-summary-p-${index + 1}`,
      text,
    }));
}

/** 按 UTF-16 上界估算一个 DOC 文档块的传输体积。 */
export function estimateDocBlockBytes(block: DocBlock) {
  // JSON 字符数按 UTF-16 上界估算，避免正文分块在中文内容下明显超出目标体积。
  return JSON.stringify(block).length * 2;
}

/** 按目标传输体积切分正文块，且不拆开单个文档块。 */
export function chunkDocBlocks(
  blocks: DocBlock[],
  targetBytes = DEFAULT_DOC_BLOCK_TARGET_BYTES,
): DocBlockChunk[] {
  const chunks: DocBlockChunk[] = [];
  const safeTargetBytes = Math.max(1, targetBytes);
  let startIndex = 0;
  let current: DocBlock[] = [];
  let currentBytes = 0;

  blocks.forEach((block, index) => {
    const blockBytes = estimateDocBlockBytes(block);
    if (current.length && currentBytes + blockBytes > safeTargetBytes) {
      chunks.push({ startIndex, blocks: current });
      startIndex = index;
      current = [];
      currentBytes = 0;
    }
    current.push(block);
    currentBytes += blockBytes;
  });

  if (current.length) chunks.push({ startIndex, blocks: current });
  return chunks;
}
