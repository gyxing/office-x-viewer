import { memo } from 'react';
import type { DocxBlock } from '../../services/docx/types';
import { DocxChartBlockView } from './DocxChartBlock';
import { DocxParagraph } from './DocxParagraph';
import { DocxTableBlockView } from './DocxTableBlock';

type DocxBlockRendererProps = {
  block: DocxBlock;
  availableWidth?: number;
};

function DocxBlockRendererComponent({ block, availableWidth }: DocxBlockRendererProps) {
  if (block.type === 'table') return <DocxTableBlockView block={block} availableWidth={availableWidth} />;
  if (block.type === 'chart') return <DocxChartBlockView block={block} zoom={100} />;
  return <DocxParagraph block={block} />;
}

export const DocxBlockRenderer = memo(DocxBlockRendererComponent);
