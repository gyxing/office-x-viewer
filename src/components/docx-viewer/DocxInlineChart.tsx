import { memo } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { OfficeChartView } from '../office-chart/OfficeChartView';

type DocxInlineChartProps = {
  inline: Extract<DocxInline, { type: 'chart' }>;
};

function DocxInlineChartComponent({ inline }: DocxInlineChartProps) {
  const chart = inline.chart;
  return (
    <span style={{ display: 'inline-block', width: chart.width, height: chart.height, verticalAlign: 'middle' }}>
      <OfficeChartView chart={chart.chart} width={chart.width} height={chart.height} zoom={100} />
    </span>
  );
}

export const DocxInlineChart = memo(DocxInlineChartComponent);
