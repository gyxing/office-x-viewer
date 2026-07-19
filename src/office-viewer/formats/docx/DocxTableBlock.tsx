// DocxTableBlock 渲染 DOCX 表格块，包括列宽、单元格边框、内边距和嵌套段落/图表。
import React, { memo } from 'react';
import type { DocxTableBlock as DocxTableBlockModel } from '../../services/docx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
import { DocxParagraph } from './DocxParagraph';
import { calculatePositionStyle } from './positionUtils';

type DocxTableBlockProps = {
  block: DocxTableBlockModel;
  availableWidth?: number;
};

function DocxTableBlockComponent({
  block,
  availableWidth,
}: DocxTableBlockProps) {
  const defaultVerticalPadding = block.insideShape ? 2 : 0;
  const resolveVerticalPadding = (value: number | undefined) =>
    block.insideShape
      ? Math.max(value ?? 0, defaultVerticalPadding)
      : value ?? 0;
  const marginLeft =
    block.align === 'center' ? 'auto' : block.align === 'right' ? 'auto' : 0;
  const marginRight =
    block.align === 'center' ? 'auto' : block.align === 'right' ? 0 : 'auto';
  const totalColumns =
    block.columns?.reduce((sum, width) => sum + width, 0) ?? block.width ?? 0;
  const shouldFit = Boolean(
    !block.position &&
      availableWidth &&
      block.width &&
      block.width > availableWidth * 1.1,
  );
  const tableWidth = block.position
    ? block.width ?? availableWidth ?? '100%'
    : shouldFit
    ? '100%'
    : block.width ?? availableWidth ?? '100%';
  const overflowWidth =
    !block.position &&
    availableWidth &&
    block.width &&
    block.width > availableWidth
      ? block.width - availableWidth
      : 0;
  const overflowMarginLeft =
    block.align === 'center'
      ? -overflowWidth / 2
      : block.align === 'right'
      ? -overflowWidth
      : undefined;
  const positionStyle = calculatePositionStyle(block.position);

  return (
    <div
      className="oxv-docx-table-block"
      style={{
        ...positionStyle,
        position: block.position ? positionStyle.position : 'relative',
        top: block.position ? positionStyle.top : block.visualOffsetTop,
        zIndex: block.position ? positionStyle.zIndex : 1,
        marginTop: block.position ? undefined : block.marginTop,
        maxWidth: block.position ? 'none' : undefined,
      }}
    >
      <table
        className="oxv-docx-table-block__table"
        style={{
          width: tableWidth,
          marginLeft: block.position
            ? 0
            : overflowWidth
            ? overflowMarginLeft
            : marginLeft,
          marginRight: block.position ? 0 : marginRight,
        }}
      >
        {block.columns?.length ? (
          <colgroup>
            {block.columns.map((width, index) => (
              <col
                key={`${block.id}-col-${index}`}
                style={{
                  width:
                    shouldFit && totalColumns > 0
                      ? `${(width / totalColumns) * 100}%`
                      : width,
                }}
              />
            ))}
          </colgroup>
        ) : null}
        <tbody>
          {block.rows.map((row) => (
            <tr
              key={row.id}
              style={{
                height: row.height,
              }}
            >
              {row.cells.map((cell) => (
                <td
                  key={cell.id}
                  className="oxv-docx-table-block__cell"
                  colSpan={
                    cell.colSpan && cell.colSpan > 1 ? cell.colSpan : undefined
                  }
                  rowSpan={
                    cell.rowSpan && cell.rowSpan > 1 ? cell.rowSpan : undefined
                  }
                  style={{
                    borderTop:
                      cell.borderTop ??
                      (cell.hasBorderTop ? 'none' : '1px solid #cfd7e3'),
                    borderRight:
                      cell.borderRight ??
                      (cell.hasBorderRight ? 'none' : '1px solid #cfd7e3'),
                    borderBottom:
                      cell.borderBottom ??
                      (cell.hasBorderBottom ? 'none' : '1px solid #cfd7e3'),
                    borderLeft:
                      cell.borderLeft ??
                      (cell.hasBorderLeft ? 'none' : '1px solid #cfd7e3'),
                    paddingTop: resolveVerticalPadding(cell.paddingTop),
                    paddingRight: cell.paddingRight ?? 7,
                    paddingBottom: resolveVerticalPadding(cell.paddingBottom),
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
                      <div
                        key={item.id}
                        className="oxv-docx-table-block__chart"
                      >
                        <OfficeChartView
                          chart={item.chart}
                          width={item.width}
                          height={item.height}
                          zoom={100}
                        />
                      </div>
                    ) : item.type === 'table' ? (
                      <DocxTableBlockComponent
                        key={item.id}
                        block={item}
                        availableWidth={cell.width ?? availableWidth}
                      />
                    ) : (
                      <DocxParagraph key={item.id} block={item} compact asDiv />
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
