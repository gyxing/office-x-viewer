import type { UnsupportedElement } from '../../../services/pptx/types';

type UnsupportedRendererProps = {
  element: UnsupportedElement;
};

export function UnsupportedRenderer({ element }: UnsupportedRendererProps) {
  return (
    <div
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
        border: '1px dashed #d92d20',
        color: '#d92d20',
        fontSize: 12,
      }}
    >
      {element.reason}
    </div>
  );
}
