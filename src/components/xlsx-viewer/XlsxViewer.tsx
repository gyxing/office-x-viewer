import { Tabs, Typography } from 'antd';
import type { CSSProperties } from 'react';
import { memo, useMemo } from 'react';
import type { XlsxCell, XlsxCellStyle, XlsxSheet, XlsxWorkbook } from '../../services/xlsx/types';
import { OfficeChartView } from '../office-chart/OfficeChartView';
import { OfficeEmpty } from '../office-viewer/OfficeEmpty';

type XlsxViewerProps = {
  workbook?: XlsxWorkbook;
  activeSheetId?: string;
  zoom: number;
  onSelectSheet: (sheetId: string) => void;
};

const EMPTY_TABS: Array<{ key: string; label: string }> = [];

function styleFromCell(cell: XlsxCell): CSSProperties {
  const style = cell.style ?? {};
  const css: CSSProperties = {
    fontWeight: style.bold ? 700 : 400,
    fontStyle: style.italic ? 'italic' : undefined,
    textDecoration: style.underline ? 'underline' : undefined,
    color: style.color,
    background: style.backgroundColor,
    textAlign: style.horizontalAlign,
    verticalAlign: style.verticalAlign,
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    whiteSpace: style.wrapText ? 'pre-wrap' : 'nowrap',
    overflowWrap: style.wrapText ? 'anywhere' : undefined,
    wordBreak: style.wrapText ? 'break-word' : undefined,
    borderColor: style.borderColor ?? (style.border ? '#b9c2d0' : '#d9e0ea'),
    borderStyle: style.border ? 'solid' : undefined,
    borderWidth: style.borderWidth ? `${style.borderWidth}px` : undefined,
  };
  return Object.fromEntries(Object.entries(css).filter(([, value]) => value !== undefined)) as CSSProperties;
}

function isInstructionCell(style?: XlsxCellStyle) {
  return Boolean(style?.color?.toLowerCase() === '#ff0000' || style?.bold);
}

function XlsxSheetGridComponent({ sheet, zoom }: { sheet: XlsxSheet; zoom: number }) {
  const scale = zoom / 100;
  const metrics = useMemo(() => {
    const tableWidth = 48 + sheet.columns.reduce((sum, column) => sum + column.width, 0);
    const tableHeight = 28 + sheet.rows.reduce((sum, row) => sum + row.height, 0);
    return {
      tableWidth,
      tableHeight,
      canvasWidth: Math.max(
        tableWidth,
        ...sheet.images.map((image) => 48 + image.x + image.width),
        ...sheet.charts.map((chart) => 48 + chart.x + chart.width),
      ),
      canvasHeight: Math.max(
        tableHeight,
        ...sheet.images.map((image) => 28 + image.y + image.height),
        ...sheet.charts.map((chart) => 28 + chart.y + chart.height),
      ),
    };
  }, [sheet]);

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

const XlsxSheetGrid = memo(XlsxSheetGridComponent);

export function XlsxViewer({ workbook, activeSheetId, zoom, onSelectSheet }: XlsxViewerProps) {
  const activeSheet = useMemo(
    () => workbook?.sheets.find((sheet) => sheet.id === activeSheetId) ?? workbook?.sheets[0],
    [activeSheetId, workbook],
  );
  const tabItems = useMemo(
    () =>
      workbook?.sheets.map((sheet) => ({
        key: sheet.id,
        label: sheet.name,
      })) ?? EMPTY_TABS,
    [workbook],
  );

  if (!activeSheet) {
    return <OfficeEmpty kind="xlsx" />;
  }

  return (
    <div
      style={{
        minHeight: 'calc(100vh - 56px)',
        height: 'calc(100vh - 56px)',
        display: 'flex',
        flexDirection: 'column',
        background: '#f2f5f9',
        overflow: 'hidden',
      }}
    >
      <div
        style={{
          flex: '0 0 auto',
          background: '#fff',
          borderBottom: '1px solid #dde3ec',
          padding: '0 16px',
          boxShadow: '0 1px 0 rgba(15, 23, 42, 0.04)',
        }}
      >
        <Tabs
          activeKey={activeSheet.id}
          onChange={onSelectSheet}
          items={tabItems}
          tabBarExtraContent={
            <Typography.Text type="secondary" style={{ fontSize: 12 }}>
              {activeSheet.range ?? `${activeSheet.rowCount} 行 x ${activeSheet.columnCount} 列`}
            </Typography.Text>
          }
        />
      </div>
      <XlsxSheetGrid sheet={activeSheet} zoom={zoom} />
    </div>
  );
}

