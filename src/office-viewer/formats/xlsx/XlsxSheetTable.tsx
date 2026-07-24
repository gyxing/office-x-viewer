// XlsxSheetTable 将工作表行列和单元格模型渲染为带表头的 HTML 表格。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { XlsxSheet } from '../../services/xlsx/types';
import { buildXlsxCellStyle, isHighlightedXlsxCell } from './sheetRenderUtils';

type XlsxSheetTableProps = {
  sheet: XlsxSheet;
  tableWidth: number;
};

function XlsxSheetTableComponent({ sheet, tableWidth }: XlsxSheetTableProps) {
  const cellStyleCache = useMemo(() => {
    const cache = new Map<string, CSSProperties>();
    sheet.rows.forEach((row) => {
      row.cells.forEach((cell) => {
        if (cell.hiddenByMerge) return;
        const important = isHighlightedXlsxCell(cell.style);
        // 大表格渲染时单元格很多，先按 ref 缓存静态样式，避免每次 JSX 展开都重复计算。
        cache.set(cell.ref, {
          height: row.height,
          minHeight: row.height,
          fontSize: important ? 14 : 13,
          ...buildXlsxCellStyle(cell),
        });
      });
    });
    return cache;
  }, [sheet]);

  return (
    <table className="oxv-xlsx-sheet-table" style={{ width: tableWidth }}>
      <colgroup>
        <col className="oxv-xlsx-sheet-table__row-header-col" />
        {sheet.columns.map((column) => (
          <col
            key={column.index}
            style={{
              width: column.width,
              display: column.hidden ? 'none' : undefined,
            }}
          />
        ))}
      </colgroup>
      <thead>
        <tr>
          <th className="oxv-xlsx-sheet-table__corner" />
          {sheet.columns.map((column) => (
            <th
              key={column.index}
              className="oxv-xlsx-sheet-table__column-header"
              style={{ display: column.hidden ? 'none' : undefined }}
            >
              {column.label}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {sheet.rows.map((row) => (
          <tr
            key={row.index}
            style={{
              height: row.height,
              display: row.hidden ? 'none' : undefined,
            }}
          >
            <th className="oxv-xlsx-sheet-table__row-header">{row.index}</th>
            {row.cells.map((cell) => {
              if (cell.hiddenByMerge) return null;
              const style = cell.style ?? {};
              const fallbackBorder = style.border
                ? `${style.borderWidth ?? 1}px solid ${
                    style.borderColor ?? '#b9c2d0'
                  }`
                : '1px solid #d9e0ea';
              return (
                <td
                  key={cell.ref}
                  className="oxv-xlsx-sheet-table__cell"
                  colSpan={cell.colSpan}
                  rowSpan={cell.rowSpan}
                  title={cell.value}
                  style={{
                    ...cellStyleCache.get(cell.ref),
                    borderTop: style.borderTop ?? fallbackBorder,
                    borderRight: style.borderRight ?? fallbackBorder,
                    borderBottom: style.borderBottom ?? fallbackBorder,
                    borderLeft: style.borderLeft ?? fallbackBorder,
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
  );
}

export const XlsxSheetTable = memo(XlsxSheetTableComponent);
