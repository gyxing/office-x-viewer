// TableRenderer 渲染 PPTX 表格元素，包括单元格填充、边框和文字样式。
import { memo } from 'react';
import type { TableElement } from '../../../services/pptx/types';

type TableRendererProps = {
  element: TableElement;
};

function colorWithOpacity(color?: string, opacity?: number) {
  if (!color || opacity === undefined || opacity >= 1) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const value = Number.parseInt(normalized, 16);
  const r = (value >> 16) & 255;
  const g = (value >> 8) & 255;
  const b = value & 255;
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

function TableRendererComponent({ element }: TableRendererProps) {
  const columnWidths = element.columnWidths ?? [];
  const rowHeights = element.rowHeights ?? [];
  return (
    <table
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
        borderCollapse: 'collapse',
        tableLayout: 'fixed',
        color: '#172033',
        fontFamily: 'inherit',
        fontSize: 12,
        background: 'transparent',
      }}
    >
      <tbody>
        {element.rows.map((row, rowIndex) => (
          <tr
            key={rowIndex}
            style={{
              height: rowHeights[rowIndex],
            }}
          >
            {row.map((cell, cellIndex) => (
              <td
                key={cellIndex}
                style={{
                  width: columnWidths[cellIndex],
                  background: colorWithOpacity(cell.backgroundColor ?? undefined, cell.backgroundOpacity),
                  borderStyle: cell.borderColor ? 'solid' : 'none',
                  borderColor: colorWithOpacity(cell.borderColor ?? undefined, cell.borderOpacity) ?? 'transparent',
                  borderWidth: cell.borderWidth ?? 1,
                  padding: `${cell.margins?.top ?? 0}px ${cell.margins?.right ?? 0}px ${cell.margins?.bottom ?? 0}px ${cell.margins?.left ?? 0}px`,
                  verticalAlign: cell.verticalAlign ?? 'middle',
                  overflow: 'hidden',
                  whiteSpace: 'pre-wrap',
                  wordBreak: 'break-word',
                }}
              >
                {cell.paragraphs?.length ? (
                  cell.paragraphs.map((paragraph, paragraphIndex) => (
                    <div
                      key={paragraphIndex}
                      style={{
                        textAlign: paragraph.style?.align ?? cell.style?.align ?? 'left',
                        lineHeight: paragraph.style?.lineHeight ?? cell.style?.lineHeight ?? 1.2,
                        whiteSpace: 'inherit',
                      }}
                    >
                      {paragraph.runs.map((run, runIndex) => (
                        <span
                          key={runIndex}
                          style={{
                            color: run.style?.color ?? cell.style?.color,
                            fontFamily: run.style?.fontFamily ?? cell.style?.fontFamily,
                            fontSize: run.style?.fontSize ?? cell.style?.fontSize,
                            fontWeight: run.style?.bold || cell.style?.bold ? 600 : 400,
                            fontStyle: run.style?.italic || cell.style?.italic ? 'italic' : 'normal',
                            textDecoration:
                              [
                                run.style?.underline || cell.style?.underline ? 'underline' : '',
                                run.style?.strike && run.style.strike !== 'none' ? 'line-through' : '',
                              ]
                                .filter(Boolean)
                                .join(' ') || 'none',
                            letterSpacing: run.style?.charSpace ?? cell.style?.charSpace ?? 0,
                          }}
                        >
                          {run.text}
                        </span>
                      ))}
                    </div>
                  ))
                ) : (
                  cell.text
                )}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export const TableRenderer = memo(TableRendererComponent);
