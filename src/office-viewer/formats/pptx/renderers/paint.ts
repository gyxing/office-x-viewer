// paint 工具负责把 PPTX 颜色、透明度和渐变模型转换成 CSS/SVG 可用格式。
import type { GradientFill } from '../../../services/pptx/types';

export function colorWithOpacity(color?: string, opacity?: number) {
  if (!color || opacity === undefined || opacity >= 1) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const value = Number.parseInt(normalized, 16);
  const r = (value >> 16) & 255;
  const g = (value >> 8) & 255;
  const b = value & 255;
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

export function isGradientPaint(paint?: string | GradientFill | null): paint is GradientFill {
  return Boolean(paint && typeof paint === 'object' && paint.type === 'linear');
}

function normalizeCssAngle(angle: number) {
  return ((angle + 90) % 360 + 360) % 360;
}

function formatOffset(offset: number) {
  return `${Math.max(0, Math.min(100, offset * 100)).toFixed(1).replace(/\.0$/, '')}%`;
}

export function paintToCss(paint?: string | GradientFill | null, opacity?: number) {
  if (!paint) return undefined;
  if (!isGradientPaint(paint)) return colorWithOpacity(paint, opacity);
  const stops = paint.stops
    .slice()
    .sort((a, b) => a.offset - b.offset)
    .map((stop) => `${stop.color} ${formatOffset(stop.offset)}`);
  return `linear-gradient(${normalizeCssAngle(paint.angle)}deg, ${stops.join(', ')})`;
}

export function gradientToSvgEndpoints(angle: number) {
  const radians = (angle * Math.PI) / 180;
  return {
    x1: 0.5 - Math.cos(radians) / 2,
    y1: 0.5 - Math.sin(radians) / 2,
    x2: 0.5 + Math.cos(radians) / 2,
    y2: 0.5 + Math.sin(radians) / 2,
  };
}
