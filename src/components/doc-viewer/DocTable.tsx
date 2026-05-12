import { memo } from 'react';
import type { DocTableBlock } from '../../services/doc/types';
import { DocInlineContent } from './DocInlineContent';
import { docTextStyleToCss } from './shared';

type DocTableProps = {
  block: DocTableBlock;
};

function DocTableComponent({ block }: DocTableProps) {
  const columnCount = Math.max(...block.rows.map((row) => row.cells.length), 1);
  const borderColor = block.style?.borderColor ?? '#cfd7e3';
  const totalColumns = block.columns?.reduce((sum, width) => sum + width, 0) ?? 0;

  return (
    <div className="oxv-doc-table">
      <table className="oxv-doc-table__table" style={{ width: block.width ?? '100%' }}>
        {block.columns?.length ? (
          <colgroup>
            {block.columns.map((width, index) => (
              <col key={`${block.id}-col-${index}`} style={{ width: totalColumns ? `${(width / totalColumns) * 100}%` : width }} />
            ))}
          </colgroup>
        ) : null}
        <tbody>
          {block.rows.map((row) => (
            <tr key={row.id}>
              {row.cells.map((cell) => (
                <td
                  key={cell.id}
                  className="oxv-doc-table__cell"
                  colSpan={cell.colSpan && cell.colSpan > 1 ? cell.colSpan : undefined}
                  style={{
                    borderTop: cell.borderTop ?? `1px solid ${borderColor}`,
                    borderRight: cell.borderRight ?? `1px solid ${borderColor}`,
                    borderBottom: cell.borderBottom ?? `1px solid ${borderColor}`,
                    borderLeft: cell.borderLeft ?? `1px solid ${borderColor}`,
                    width: cell.width,
                    verticalAlign: cell.verticalAlign ?? 'top',
                    ...docTextStyleToCss(cell.style),
                  }}
                >
                  <DocInlineContent inlines={cell.inlines} fallback={cell.text} />
                </td>
              ))}
              {row.cells.length < columnCount
                ? Array.from({ length: columnCount - row.cells.length }).map((_, index) => (
                    <td key={`${row.id}-empty-${index}`} className="oxv-doc-table__cell" style={{ border: `1px solid ${borderColor}` }} />
                  ))
                : null}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export const DocTable = memo(DocTableComponent);
