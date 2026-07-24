import type { VectorElement, VectorScene, VectorStyle } from './types';

function escapeXml(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function styleAttributes(style: VectorStyle, isText = false) {
  const fill = isText ? style.textColor ?? style.fill ?? '#000000' : style.fill;
  return [
    `stroke="${escapeXml(style.stroke ?? 'none')}"`,
    `fill="${escapeXml(fill ?? 'none')}"`,
    `stroke-width="${style.strokeWidth ?? 1}"`,
    style.opacity === undefined ? '' : `opacity="${style.opacity}"`,
    style.fontFamily ? `font-family="${escapeXml(style.fontFamily)}"` : '',
    style.fontSize ? `font-size="${style.fontSize}"` : '',
    style.fontWeight ? `font-weight="${style.fontWeight}"` : '',
  ]
    .filter(Boolean)
    .join(' ');
}

function serializeElement(element: VectorElement) {
  const style = styleAttributes(element.style, element.type === 'text');
  if (element.type === 'line') {
    return `<line x1="${element.x1}" y1="${element.y1}" x2="${element.x2}" y2="${element.y2}" ${style}/>`;
  }
  if (element.type === 'polyline' || element.type === 'polygon') {
    const points = element.points.map(([x, y]) => `${x},${y}`).join(' ');
    return `<${element.type} points="${points}" ${style}/>`;
  }
  if (element.type === 'rectangle') {
    return `<rect x="${element.x}" y="${element.y}" width="${
      element.width
    }" height="${element.height}" rx="${element.radiusX ?? 0}" ry="${
      element.radiusY ?? 0
    }" ${style}/>`;
  }
  if (element.type === 'ellipse') {
    return `<ellipse cx="${element.x + element.width / 2}" cy="${
      element.y + element.height / 2
    }" rx="${element.width / 2}" ry="${element.height / 2}" ${style}/>`;
  }
  if (element.type === 'path') {
    return `<path d="${escapeXml(element.data)}" ${style}/>`;
  }
  if (element.type === 'text') {
    return `<text x="${element.x}" y="${element.y}" ${style}>${escapeXml(
      element.text,
    )}</text>`;
  }
  if (element.type === 'image') {
    const safeDataUrl = element.dataUrl.startsWith('data:')
      ? escapeXml(element.dataUrl)
      : '';
    return `<image x="${element.x}" y="${element.y}" width="${element.width}" height="${element.height}" href="${safeDataUrl}" ${style}/>`;
  }
  return '';
}

/** 将内部矢量场景安全序列化为独立 SVG 字符串。 */
export function vectorSceneToSvg(scene: VectorScene) {
  const [x, y, width, height] = scene.viewBox;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${
    scene.width
  }" height="${
    scene.height
  }" viewBox="${x} ${y} ${width} ${height}">${scene.elements
    .map(serializeElement)
    .join('')}</svg>`;
}
