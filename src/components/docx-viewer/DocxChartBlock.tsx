// DocxChartBlock 渲染 DOCX 文档中的独立图表块。
import { memo } from 'react';
import type { DocxChartBlock as DocxChartBlockModel } from '../../services/docx/types';
import { OfficeChartView } from '../office-chart/OfficeChartView';

type DocxChartBlockProps = {
  block: DocxChartBlockModel;
  zoom: number;
};

function DocxChartBlockComponent({ block, zoom }: DocxChartBlockProps) {
  return (
    <div className="oxv-docx-chart-block">
      <OfficeChartView chart={block.chart} width={block.width} height={block.height} zoom={zoom} />
    </div>
  );
}

export const DocxChartBlock = memo(DocxChartBlockComponent);
