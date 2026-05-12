import { memo } from 'react';
import type { DocBlock } from '../../services/doc/types';
import { DocList } from './DocList';
import { DocParagraph } from './DocParagraph';
import { DocTable } from './DocTable';

type DocBlockRendererProps = {
  block: DocBlock;
};

function DocBlockRendererComponent({ block }: DocBlockRendererProps) {
  if (block.type === 'table') return <DocTable block={block} />;
  if (block.type === 'list') return <DocList block={block} />;
  return <DocParagraph block={block} />;
}

export const DocBlockRenderer = memo(DocBlockRendererComponent);
