import type { DocxPosition } from '../../services/docx/types';

const CSS_MAX_Z_INDEX = 2147483647;

/**
 * 安全地格式化数值为像素字符串
 */
function safePx(value: number | undefined): string {
  if (value === undefined || !Number.isFinite(value)) {
    return '0px';
  }
  return `${Math.round(value)}px`;
}

/**
 * 计算元素的定位样式
 * 根据不同的参考系（page、margin、column等）计算正确的 left 和 top 值
 */
export function calculatePositionStyle(position: DocxPosition | undefined) {
  if (!position) {
    return {
      position: undefined,
      left: undefined,
      top: undefined,
      zIndex: undefined,
      transform: undefined,
      transformOrigin: undefined,
    };
  }

  const { left, top, relativeFromH, relativeFromV, zIndex, behindDoc, rotation, flipH, flipV } = position;

  // 确保 left 和 top 是有效数值
  const safeLeft = Number.isFinite(left) ? left : 0;
  const safeTop = Number.isFinite(top) ? top : 0;

  // 水平定位计算
  let calculatedLeft: string | number = safePx(safeLeft);
  if (relativeFromH === 'page') {
    // 相对于页面：需要减去页面左边距
    calculatedLeft = `calc(${safePx(safeLeft)} - var(--oxv-docx-page-margin-left, 0px))`;
  } else if (relativeFromH === 'margin') {
    // 相对于边距区域：从内容区域左边缘开始
    calculatedLeft = safePx(safeLeft);
  } else if (relativeFromH === 'column') {
    // 相对于列：从当前列左边缘开始
    calculatedLeft = safePx(safeLeft);
  } else if (relativeFromH === 'leftMargin') {
    // 相对于左边距：从页面左边缘开始，减去左边距
    calculatedLeft = `calc(${safePx(safeLeft)} - var(--oxv-docx-page-margin-left, 0px))`;
  } else if (relativeFromH === 'rightMargin') {
    // 相对于右边距：从页面右边缘开始
    calculatedLeft = `calc(var(--oxv-docx-page-width, 100%) - var(--oxv-docx-page-margin-right, 0px) + ${safePx(safeLeft)} - var(--oxv-docx-page-margin-left, 0px))`;
  } else if (relativeFromH === 'insideMargin') {
    // 内侧边距（奇数页=左，偶数页=右）
    calculatedLeft = safePx(safeLeft);
  } else if (relativeFromH === 'outsideMargin') {
    // 外侧边距（奇数页=右，偶数页=左）
    calculatedLeft = safePx(safeLeft);
  } else if (relativeFromH === 'character') {
    // 相对于字符位置
    calculatedLeft = safePx(safeLeft);
  }

  // 垂直定位计算
  let calculatedTop: string | number = safePx(safeTop);
  if (relativeFromV === 'page') {
    // 相对于页面：需要减去页面上边距
    calculatedTop = `calc(${safePx(safeTop)} - var(--oxv-docx-page-margin-top, 0px))`;
  } else if (relativeFromV === 'margin') {
    // 相对于边距区域：从内容区域顶部开始
    calculatedTop = safePx(safeTop);
  } else if (relativeFromV === 'paragraph') {
    // 相对于段落：从当前段落顶部开始
    calculatedTop = safePx(safeTop);
  } else if (relativeFromV === 'line') {
    // 相对于行：从当前行顶部开始
    calculatedTop = safePx(safeTop);
  } else if (relativeFromV === 'text') {
    calculatedTop = `calc(var(--oxv-docx-page-margin-top, 0px) + ${safePx(safeTop)})`;
  } else if (relativeFromV === 'topMargin') {
    // 相对于上边距：从页面顶部开始，减去上边距
    calculatedTop = `calc(${safePx(safeTop)} - var(--oxv-docx-page-margin-top, 0px))`;
  } else if (relativeFromV === 'bottomMargin') {
    // 相对于下边距：从页面底部开始
    calculatedTop = `calc(var(--oxv-docx-page-height, 100%) - var(--oxv-docx-page-margin-bottom, 0px) + ${safePx(safeTop)} - var(--oxv-docx-page-margin-top, 0px))`;
  } else if (relativeFromV === 'insideMargin') {
    // 内侧边距（用于双面打印）
    calculatedTop = safePx(safeTop);
  } else if (relativeFromV === 'outsideMargin') {
    // 外侧边距（用于双面打印）
    calculatedTop = safePx(safeTop);
  }

  // 计算 transform
  const transforms: string[] = [];
  if (rotation && Number.isFinite(rotation) && rotation !== 0) {
    // 限制旋转角度在合理范围内
    const normalizedRotation = ((rotation % 360) + 360) % 360;
    transforms.push(`rotate(${normalizedRotation}deg)`);
  }
  if (flipH) {
    transforms.push('scaleX(-1)');
  }
  if (flipV) {
    transforms.push('scaleY(-1)');
  }

  // 计算 z-index：behindDoc 元素应在纸张之上、正文之下；负数会被页面白底盖住。
  let calculatedZIndex: number | undefined;
  if (behindDoc) {
    calculatedZIndex = 0;
  } else if (zIndex !== undefined && Number.isFinite(zIndex)) {
    // OOXML 的 relativeHeight 本身表达层叠顺序，压缩会抹平相邻对象的前后关系。
    calculatedZIndex = Math.min(CSS_MAX_Z_INDEX, Math.max(1, Math.round(zIndex)));
  } else {
    // 未指定 z-index 的浮动对象默认放在正文之上，但低于明确指定前景层级的对象。
    calculatedZIndex = 2;
  }

  return {
    position: 'absolute' as const,
    left: calculatedLeft,
    top: calculatedTop,
    zIndex: calculatedZIndex,
    transform: transforms.length > 0 ? transforms.join(' ') : undefined,
    transformOrigin: transforms.length > 0 ? 'center center' : undefined,
  };
}

/**
 * 判断元素是否有定位
 */
export function hasPosition(position: DocxPosition | undefined): boolean {
  return Boolean(position);
}

/**
 * 判断元素是否在文档流中（非定位元素）
 */
export function isInFlow(position: DocxPosition | undefined): boolean {
  return !hasPosition(position);
}
