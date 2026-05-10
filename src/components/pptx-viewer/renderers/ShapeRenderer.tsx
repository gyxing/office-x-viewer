import type { ShapeElement } from '../../../services/pptx/types';

type ShapeRendererProps = {
  element: ShapeElement;
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

function lineStyle(dash?: string) {
  if (!dash || dash === 'solid') return 'solid';
  if (dash.includes('dot')) return 'dotted';
  return 'dashed';
}

export function ShapeRenderer({ element }: ShapeRendererProps) {
  const radius = element.shape === 'roundRect' ? Math.min(element.width, element.height) * (element.borderRadius ?? 0.12) : 0;
  const shadow = element.shadow
    ? `${element.shadow.offsetX ?? 0}px ${element.shadow.offsetY ?? 0}px ${Math.max(0, element.shadow.blur ?? 0)}px ${colorWithOpacity(element.shadow.color ?? 'rgba(0,0,0,0.18)', element.shadow.opacity)}`
    : undefined;

  return (
    <div
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
        background: element.fill
          ? element.fillOpacity !== undefined
            ? colorWithOpacity(element.fill, element.fillOpacity)
            : element.fill
          : 'transparent',
        border: element.stroke
          ? `${element.strokeWidth ?? 1}px ${lineStyle(element.strokeDash)} ${colorWithOpacity(element.stroke, element.strokeOpacity)}`
          : undefined,
        borderRadius: radius,
        boxShadow: shadow,
        transform: element.rotate
          ? `rotate(${element.rotate}deg)${element.flipH ? ' scaleX(-1)' : ''}${element.flipV ? ' scaleY(-1)' : ''}`
          : `${element.flipH ? 'scaleX(-1)' : ''}${element.flipV ? ' scaleY(-1)' : ''}`.trim(),
        transformOrigin: 'center center',
      }}
    />
  );
}
