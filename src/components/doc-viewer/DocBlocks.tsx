import { memo, useMemo } from 'react';
import type { ReactNode } from 'react';
import type { DocBlock } from '../../services/doc/types';
import { DocBlockRenderer } from './DocBlockRenderer';
import { DocImageLayout } from './DocImageLayout';
import { imagesFromImageOnlyParagraph } from './shared';

type DocBlocksProps = {
  blocks: DocBlock[];
  contentWidth: number;
};

function buildDocBlocks(blocks: DocBlock[], contentWidth: number) {
  const renderedBlocks: ReactNode[] = [];
  let index = 0;

  while (index < blocks.length) {
    const images = imagesFromImageOnlyParagraph(blocks[index]);
    if (!images.length) {
      renderedBlocks.push(<DocBlockRenderer key={blocks[index].id} block={blocks[index]} />);
      index += 1;
      continue;
    }

    const imageGroup = [...images];
    let nextIndex = index + 1;
    while (nextIndex < blocks.length) {
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

function DocBlocksComponent({ blocks, contentWidth }: DocBlocksProps) {
  const renderedBlocks = useMemo(() => buildDocBlocks(blocks, contentWidth), [blocks, contentWidth]);
  return <>{renderedBlocks}</>;
}

export const DocBlocks = memo(DocBlocksComponent);
