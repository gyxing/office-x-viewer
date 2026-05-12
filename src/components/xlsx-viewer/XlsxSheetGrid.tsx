import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { XlsxSheet } from '../../services/xlsx/types';
import { getSheetMetrics } from './shared';
import { XlsxFloatingCharts } from './XlsxFloatingCharts';
import { XlsxFloatingImages } from './XlsxFloatingImages';
import { XlsxSheetTable } from './XlsxSheetTable';

type XlsxSheetGridProps = {
  sheet: XlsxSheet;
  zoom: number;
};

function XlsxSheetGridComponent({ sheet, zoom }: XlsxSheetGridProps) {
  const scale = zoom / 100;
  const metrics = useMemo(() => getSheetMetrics(sheet), [sheet]);
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
