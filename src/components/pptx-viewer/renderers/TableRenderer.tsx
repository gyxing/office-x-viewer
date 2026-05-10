import type { TableElement } from '../../../services/pptx/types';

type TableRendererProps = {
  element: TableElement;
};

export function TableRenderer({ element }: TableRendererProps) {
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
        overflow: 'hidden',
        color: '#172033',
        fontFamily: 'inherit',
        fontSize: 12,
      }}
    >
      <tbody>
        {element.rows.map((row, rowIndex) => (
          <tr key={rowIndex}>
            {row.map((cell, cellIndex) => (
              <td
                key={cellIndex}
                style={{
                  border: '1px solid rgba(221, 227, 236, 0.95)',
                  padding: '2px 6px',
                  verticalAlign: 'middle',
                  overflow: 'hidden',
                  textOverflow: 'clip',
                  whiteSpace: 'pre-wrap',
                }}
              >
                {cell.text}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
