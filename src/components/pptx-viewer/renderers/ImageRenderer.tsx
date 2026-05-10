import type { ImageElement } from '../../../services/pptx/types';

type ImageRendererProps = {
  element: ImageElement;
};

export function ImageRenderer({ element }: ImageRendererProps) {
  const left = element.crop?.left ?? 0;
  const top = element.crop?.top ?? 0;
  const right = element.crop?.right ?? 0;
  const bottom = element.crop?.bottom ?? 0;
  const visibleWidth = Math.max(0.01, 1 - left - right);
  const visibleHeight = Math.max(0.01, 1 - top - bottom);

  return (
    <div
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
        overflow: 'hidden',
        transform: [
          element.rotate ? `rotate(${element.rotate}deg)` : '',
          element.flipH ? 'scaleX(-1)' : '',
          element.flipV ? 'scaleY(-1)' : '',
        ]
          .filter(Boolean)
          .join(' '),
        transformOrigin: 'center center',
        pointerEvents: 'none',
      }}
      >
      <img
        alt={element.alt ?? ''}
        src={element.src}
        style={{
          position: 'absolute',
          left: `${-(left / visibleWidth) * 100}%`,
          top: `${-(top / visibleHeight) * 100}%`,
          width: `${100 / visibleWidth}%`,
          height: `${100 / visibleHeight}%`,
          objectFit: 'fill',
          display: 'block',
        }}
      />
    </div>
  );
}
