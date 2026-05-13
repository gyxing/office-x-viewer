// DocxTableBlock 渲染 DOCX 表格块，包括列宽、单元格边框、内边距和嵌套段落/图表。
import { memo } from 'react';
import type { DocxTableBlock as DocxTableBlockModel } from '../../services/docx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
import { DocxParagraph } from './DocxParagraph';

type DocxTableBlockProps = {
  block: DocxTableBlockModel;
  availableWidth?: number;
};

function DocxTableBlockComponent({ block, availableWidth }: DocxTableBlockProps) {
  const marginLeft = block.align === 'center' ? 'auto' : block.align === 'right' ? 'auto' : 0;
  const marginRight = block.align === 'center' ? 'auto' : block.align === 'right' ? 0 : 'auto';
  const totalColumns = block.columns?.reduce((sum, width) => sum + width, 0) ?? block.width ?? 0;
  const shouldFit = Boolean(availableWidth && block.width && block.width > availableWidth);
  const tableWidth = shouldFit ? '100%' : block.width ?? availableWidth ?? '100%';

  return (
    <div className="oxv-docx-table-block">
      <table
        className="oxv-docx-table-block__table"
        style={{
          width: tableWidth,
          marginLeft,
          marginRight,
        }}
      >
        {block.columns?.length ? (
          <colgroup>
            {block.columns.map((width, index) => (
              <col
                key={`${block.id}-col-${index}`}
                style={{
                  width: shouldFit && totalColumns > 0 ? `${(width / totalColumns) * 100}%` : width,
                }}
              />
            ))}
          </colgroup>
        ) : null}
        <tbody>
          {block.rows.map((row) => (
            <tr key={row.id}>
              {row.cells.map((cell) => (
                <td
                  key={cell.id}
                  className="oxv-docx-table-block__cell"
                  colSpan={cell.colSpan && cell.colSpan > 1 ? cell.colSpan : undefined}
                  style={{
                    borderTop: cell.borderTop ?? (cell.hasBorderTop ? 'none' : '1px solid #cfd7e3'),
                    borderRight: cell.borderRight ?? (cell.hasBorderRight ? 'none' : '1px solid #cfd7e3'),
                    borderBottom: cell.borderBottom ?? (cell.hasBorderBottom ? 'none' : '1px solid #cfd7e3'),
                    borderLeft: cell.borderLeft ?? (cell.hasBorderLeft ? 'none' : '1px solid #cfd7e3'),
                    paddingTop: cell.paddingTop ?? 0,
                    paddingRight: cell.paddingRight ?? 7,
                    paddingBottom: cell.paddingBottom ?? 0,
                    paddingLeft: cell.paddingLeft ?? 7,
                    width: shouldFit ? undefined : cell.width,
                    verticalAlign: cell.verticalAlign,
                    background: cell.backgroundColor ?? '#fff',
                    wordBreak: cell.noWrap ? 'normal' : 'break-word',
                    overflowWrap: cell.noWrap ? 'normal' : 'anywhere',
                    whiteSpace: cell.noWrap ? 'nowrap' : undefined,
                  }}
                >
                  {cell.blocks.map((item) =>
                    item.type === 'chart' ? (
                      <div key={item.id} className="oxv-docx-table-block__chart">
                        <OfficeChartView chart={item.chart} width={item.width} height={item.height} zoom={100} />
                      </div>
                    ) : (
                      <DocxParagraph key={item.id} block={item} compact />
                    ),
                  )}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export const DocxTableBlock = memo(DocxTableBlockComponent);
