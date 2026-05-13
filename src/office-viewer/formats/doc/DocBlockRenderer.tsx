// DocBlockRenderer 根据 DOC 块类型分发到段落、列表或表格渲染组件。
import { memo } from 'react';
import type { DocBlock } from '../../services/doc/types';
import { DocListBlock } from './DocListBlock';
import { DocParagraphBlock } from './DocParagraphBlock';
import { DocTableBlock } from './DocTableBlock';

type DocBlockRendererProps = {
  block: DocBlock;
};

function DocBlockRendererComponent({ block }: DocBlockRendererProps) {
  if (block.type === 'table') return <DocTableBlock block={block} />;
  if (block.type === 'list') return <DocListBlock block={block} />;
  return <DocParagraphBlock block={block} />;
}

export const DocBlockRenderer = memo(DocBlockRendererComponent);
