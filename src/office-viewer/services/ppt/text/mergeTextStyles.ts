import type { TextStyle } from '../../presentation/types';

/** 按继承顺序合并文本样式，同时保留 false、0 等显式覆盖值。 */
export function mergePptTextStyles(
  ...styles: Array<TextStyle | undefined>
): TextStyle {
  return styles.reduce<TextStyle>(
    (merged, style) => (style ? { ...merged, ...style } : merged),
    {},
  );
}
