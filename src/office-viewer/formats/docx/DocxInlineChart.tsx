// DocxInlineChart 渲染 DOCX 行内图表。
import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';

type DocxInlineChartProps = {
  inline: Extract<DocxInline, { type: 'chart' }>;
};

function DocxInlineChartComponent({ inline }: DocxInlineChartProps) {
  const chart = inline.chart;
  const chartStyle = useMemo<CSSProperties>(
    () =>
      ({
        '--oxv-docx-inline-chart-width': `${chart.width}px`,
        '--oxv-docx-inline-chart-height': `${chart.height}px`,
      }) as CSSProperties,
    [chart.height, chart.width],
  );

  return (
    <span className="oxv-docx-inline-chart" style={chartStyle}>
      <OfficeChartView chart={chart.chart} width={chart.width} height={chart.height} zoom={100} />
    </span>
  );
}

export const DocxInlineChart = memo(DocxInlineChartComponent);
