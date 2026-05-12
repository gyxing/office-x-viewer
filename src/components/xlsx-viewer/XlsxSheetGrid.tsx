// XlsxSheetGrid 负责工作表滚动画布，统一承载表格、浮动图片和浮动图表。
import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { XlsxSheet } from '../../services/xlsx/types';
import { getXlsxSheetMetrics } from './sheetRenderUtils';
import { XlsxFloatingCharts } from './XlsxFloatingCharts';
import { XlsxFloatingImages } from './XlsxFloatingImages';
import { XlsxSheetTable } from './XlsxSheetTable';

type XlsxSheetGridProps = {
  sheet: XlsxSheet;
  zoom: number;
};

function XlsxSheetGridComponent({ sheet, zoom }: XlsxSheetGridProps) {
  const scale = zoom / 100;
  const metrics = useMemo(() => getXlsxSheetMetrics(sheet), [sheet]);
  // 这里使用 zoom 是为了让表格、图片和图表保持同一个坐标系缩放。
  const canvasStyle = useMemo<CSSProperties>(
    () => ({
      width: metrics.canvasWidth,
      minWidth: metrics.canvasWidth,
      minHeight: metrics.canvasHeight,
      zoom: scale,
    }),
    [metrics.canvasHeight, metrics.canvasWidth, scale],
  );

  return (
    <div className="oxv-xlsx-sheet-grid">
      <div className="oxv-xlsx-sheet-grid__canvas" style={canvasStyle}>
        <XlsxSheetTable sheet={sheet} tableWidth={metrics.tableWidth} />
        <XlsxFloatingImages images={sheet.images} />
        <XlsxFloatingCharts charts={sheet.charts} />
      </div>
    </div>
  );
}

export const XlsxSheetGrid = memo(XlsxSheetGridComponent);
