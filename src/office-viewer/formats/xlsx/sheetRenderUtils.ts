// sheetRenderUtils 提供 XLSX 工作表渲染所需的样式转换和画布尺寸计算。
import type { CSSProperties } from 'react';
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../../services/xlsx/types';

// 单元格字体、填充、边框等来自工作簿样式表，只能在运行时转成 CSS，不能放到 Less。
export function buildXlsxCellStyle(cell: XlsxCell): CSSProperties {
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

export function isHighlightedXlsxCell(style?: XlsxCellStyle) {
  return Boolean(style?.color?.toLowerCase() === '#ff0000' || style?.bold);
}

// 画布需要同时覆盖单元格区域和浮动图片/图表，否则滚动区域会截断绝对定位元素。
export function getXlsxSheetMetrics(sheet: XlsxSheet) {
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
}
