// DocxBlockRenderer 根据 DOCX 块类型分发到段落、表格或图表渲染组件。
import { memo } from 'react';
import type { DocxBlock } from '../../services/docx/types';
import { DocxChartBlock } from './DocxChartBlock';
import { DocxParagraph } from './DocxParagraph';
import { DocxTableBlock } from './DocxTableBlock';

type DocxBlockRendererProps = {
  block: DocxBlock;
  availableWidth?: number;
};

function DocxBlockRendererComponent({ block, availableWidth }: DocxBlockRendererProps) {
  if (block.type === 'table') return <DocxTableBlock block={block} availableWidth={availableWidth} />;
  if (block.type === 'chart') return <DocxChartBlock block={block} zoom={100} />;
  return <DocxParagraph block={block} />;
}

export const DocxBlockRenderer = memo(DocxBlockRendererComponent);
