// DocContentRenderer 渲染 DOC 内容块列表，并合并连续图片段落以优化排版。
import { memo, useMemo } from 'react';
import type { ReactNode } from 'react';
import type { DocBlock } from '../../services/doc/types';
import { DocBlockRenderer } from './DocBlockRenderer';
import { DocImageLayout } from './DocImageLayout';
import { imagesFromImageOnlyParagraph } from './docRenderUtils';

type DocContentRendererProps = {
  blocks: DocBlock[];
  contentWidth: number;
};

function buildDocContent(blocks: DocBlock[], contentWidth: number) {
  const renderedBlocks: ReactNode[] = [];
  let index = 0;

  while (index < blocks.length) {
    // DOC 解析出的图片经常是连续的“纯图片段落”，合并后再排版更接近 Word 的视觉结果。
    const images = imagesFromImageOnlyParagraph(blocks[index]);
    if (!images.length) {
      renderedBlocks.push(<DocBlockRenderer key={blocks[index].id} block={blocks[index]} />);
      index += 1;
      continue;
    }

    const imageGroup = [...images];
    let nextIndex = index + 1;
    while (nextIndex < blocks.length) {
      // 连续图片段落作为一个图片组渲染，后续根据宽度决定单列或双列。
      const nextImages = imagesFromImageOnlyParagraph(blocks[nextIndex]);
      if (!nextImages.length) break;
      imageGroup.push(...nextImages);
      nextIndex += 1;
    }

    renderedBlocks.push(<DocImageLayout key={`doc-image-layout-${index}`} images={imageGroup} contentWidth={contentWidth} />);
    index = nextIndex;
  }

  return renderedBlocks;
}

function DocContentRendererComponent({ blocks, contentWidth }: DocContentRendererProps) {
  const renderedBlocks = useMemo(() => buildDocContent(blocks, contentWidth), [blocks, contentWidth]);
  return <>{renderedBlocks}</>;
}

export const DocContentRenderer = memo(DocContentRendererComponent);
