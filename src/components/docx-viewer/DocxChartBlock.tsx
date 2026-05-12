import { memo } from 'react';
import type { DocxChartBlock } from '../../services/docx/types';
import { OfficeChartView } from '../office-chart/OfficeChartView';

type DocxChartBlockProps = {
  block: DocxChartBlock;
  zoom: number;
};

function DocxChartBlockViewComponent({ block, zoom }: DocxChartBlockProps) {
  return (
    <div className="oxv-docx-chart-block">
      <OfficeChartView chart={block.chart} width={block.width} height={block.height} zoom={zoom} />
    </div>
  );
}

export const DocxChartBlockView = memo(DocxChartBlockViewComponent);
