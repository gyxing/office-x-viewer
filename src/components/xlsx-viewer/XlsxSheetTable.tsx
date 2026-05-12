import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { XlsxSheet } from '../../services/xlsx/types';
import { isInstructionCell, styleFromCell } from './shared';

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
        const important = isInstructionCell(cell.style);
        cache.set(cell.ref, {
          height: row.height,
          minHeight: row.height,
          fontSize: important ? 14 : 13,
          ...styleFromCell(cell),
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
          <col key={column.index} style={{ width: column.width, display: column.hidden ? 'none' : undefined }} />
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
          <tr key={row.index} style={{ height: row.height }}>
            <th className="oxv-xlsx-sheet-table__row-header">{row.index}</th>
            {row.cells.map((cell) => {
              if (cell.hiddenByMerge) return null;
              const style = cell.style ?? {};
              return (
                <td
                  key={cell.ref}
                  className="oxv-xlsx-sheet-table__cell"
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
  );
}

export const XlsxSheetTable = memo(XlsxSheetTableComponent);
