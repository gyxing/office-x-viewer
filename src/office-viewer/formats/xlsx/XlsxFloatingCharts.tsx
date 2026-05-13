// XlsxFloatingCharts 渲染锚定在工作表画布上的浮动图表。
import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { XlsxChart } from '../../services/xlsx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';

type XlsxFloatingChartsProps = {
  charts: XlsxChart[];
};

function XlsxFloatingChart({ chart }: { chart: XlsxChart }) {
  const chartStyle = useMemo<CSSProperties>(
    () => ({
      left: 48 + chart.x,
      top: 28 + chart.y,
      width: chart.width,
      height: chart.height,
    }),
    [chart.height, chart.width, chart.x, chart.y],
  );

  return (
    <div className="oxv-xlsx-sheet-grid__floating-chart" style={chartStyle}>
      <OfficeChartView chart={chart.chart} width={chart.width} height={chart.height} zoom={100} />
    </div>
  );
}

const MemoXlsxFloatingChart = memo(XlsxFloatingChart);

function XlsxFloatingChartsComponent({ charts }: XlsxFloatingChartsProps) {
  return (
    <>
      {charts.map((chart) => (
        <MemoXlsxFloatingChart key={chart.id} chart={chart} />
      ))}
    </>
  );
}

export const XlsxFloatingCharts = memo(XlsxFloatingChartsComponent);
