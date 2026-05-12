import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { XlsxSheet } from '../../services/xlsx/types';
import { OfficeChartView } from '../office-chart/OfficeChartView';
import { getSheetMetrics, isInstructionCell, styleFromCell } from './shared';

type XlsxSheetGridProps = {
  sheet: XlsxSheet;
  zoom: number;
};

function XlsxSheetGridComponent({ sheet, zoom }: XlsxSheetGridProps) {
  const scale = zoom / 100;
  const metrics = useMemo(() => getSheetMetrics(sheet), [sheet]);

  const cellStyleCache = useMemo(() => {
    const cache = new Map<string, CSSProperties>();
    sheet.rows.forEach((row) => {
      row.cells.forEach((cell) => {
        if (cell.hiddenByMerge) return;
        const important = isInstructionCell(cell.style);
        cache.set(cell.ref, {
          height: row.height,
          minHeight: row.height,
          padding: '0 4px',
          border: '1px solid #d9e0ea',
          color: '#1f2937',
          fontSize: important ? 14 : 13,
          lineHeight: 1.2,
          overflow: 'visible',
          ...styleFromCell(cell),
        });
      });
    });
    return cache;
  }, [sheet]);

  return (
    <div
      style={{
        flex: '1 1 auto',
        minHeight: 0,
        minWidth: 0,
        overflow: 'auto',
        padding: 16,
        background: '#e9edf3',
        scrollbarGutter: 'stable both-edges',
      }}
    >
      <div
        style={{
          position: 'relative',
          width: metrics.canvasWidth,
          minWidth: metrics.canvasWidth,
          minHeight: metrics.canvasHeight,
          zoom: scale,
        }}
      >
        <table
          style={{
            borderCollapse: 'collapse',
            tableLayout: 'fixed',
            background: '#fff',
            boxShadow: '0 10px 24px rgba(15, 23, 42, 0.10)',
            border: '1px solid #b9c4d2',
            width: metrics.tableWidth,
          }}
        >
          <colgroup>
            <col style={{ width: 48 }} />
            {sheet.columns.map((column) => (
              <col key={column.index} style={{ width: column.width, display: column.hidden ? 'none' : undefined }} />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th
                style={{
                  height: 28,
                  background: '#e5ebf2',
                  border: '1px solid #c5cfdc',
                  color: '#536174',
                  fontSize: 12,
                  fontWeight: 600,
                  position: 'sticky',
                  top: 0,
                  zIndex: 3,
                }}
              />
              {sheet.columns.map((column) => (
                <th
                  key={column.index}
                  style={{
                    height: 28,
                    background: '#e5ebf2',
                    border: '1px solid #c5cfdc',
                    color: '#536174',
                    fontSize: 12,
                    fontWeight: 600,
                    position: 'sticky',
                    top: 0,
                    zIndex: 3,
                    display: column.hidden ? 'none' : undefined,
                  }}
                >
                  {column.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sheet.rows.map((row) => (
              <tr key={row.index} style={{ height: row.height }}>
                <th
                  style={{
                    width: 48,
                    background: '#e5ebf2',
                    border: '1px solid #c5cfdc',
                    color: '#536174',
                    fontSize: 12,
                    fontWeight: 600,
                    position: 'sticky',
                    left: 0,
                    zIndex: 2,
                  }}
                >
                  {row.index}
                </th>
                {row.cells.map((cell) => {
                  if (cell.hiddenByMerge) return null;
                  const style = cell.style ?? {};
                  return (
                    <td
                      key={cell.ref}
                      colSpan={cell.colSpan}
                      rowSpan={cell.rowSpan}
                      title={cell.value}
                      style={{
                        ...cellStyleCache.get(cell.ref),
                        borderTop: style.borderTop ?? '1px solid #d9e0ea',
                        borderRight: style.borderRight ?? '1px solid #d9e0ea',
                        borderBottom: style.borderBottom ?? '1px solid #d9e0ea',
                        borderLeft: style.borderLeft ?? '1px solid #d9e0ea',
                      }}
                    >
                      {cell.value}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
        {sheet.images.map((image) => (
          <img
            key={image.id}
            src={image.src}
            alt={image.alt ?? ''}
            title={image.name}
            style={{
              position: 'absolute',
              left: 48 + image.x,
              top: 28 + image.y,
              width: image.width,
              height: image.height,
              objectFit: 'contain',
              pointerEvents: 'none',
              zIndex: 4,
            }}
          />
        ))}
        {sheet.charts.map((chart) => (
          <div
            key={chart.id}
            style={{
              position: 'absolute',
              left: 48 + chart.x,
              top: 28 + chart.y,
              width: chart.width,
              height: chart.height,
              zIndex: 5,
            }}
          >
            <OfficeChartView chart={chart.chart} width={chart.width} height={chart.height} zoom={100} />
          </div>
        ))}
      </div>
    </div>
  );
}

export const XlsxSheetGrid = memo(XlsxSheetGridComponent);

