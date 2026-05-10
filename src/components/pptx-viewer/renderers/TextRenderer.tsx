import type { TextElement } from '../../../services/pptx/types';

type TextRendererProps = {
  element: TextElement;
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

function shadowToCss(element: TextElement) {
  if (!element.shadow) return undefined;
  return `${element.shadow.offsetX ?? 0}px ${element.shadow.offsetY ?? 0}px ${Math.max(0, element.shadow.blur ?? 0)}px ${colorWithOpacity(element.shadow.color ?? '#000000', element.shadow.opacity ?? 0.18)}`;
}

function radiusToPx(element: TextElement) {
  if (element.shape !== 'roundRect') return 0;
  const ratio = element.borderRadius ?? 0.12;
  return Math.min(element.width, element.height) * ratio;
}

function textDecoration(style: NonNullable<TextElement['boxStyle']>) {
  const parts: string[] = [];
  if (style.underline) parts.push('underline');
  if (style.strike && style.strike !== 'none') parts.push('line-through');
  return parts.length ? parts.join(' ') : 'none';
}

function lineStyle(dash?: string) {
  if (!dash || dash === 'solid') return 'solid';
  if (dash.includes('dot')) return 'dotted';
  return 'dashed';
}

export function TextRenderer({ element }: TextRendererProps) {
  const style = element.boxStyle ?? {};
  return (
    <div
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
        color: colorWithOpacity(style.color ?? '#172033', style.opacity),
        fontFamily: style.fontFamily,
        fontSize: style.fontSize,
        fontWeight: style.bold ? 600 : 400,
        fontStyle: style.italic ? 'italic' : 'normal',
        textDecoration: textDecoration(style),
        textTransform: style.allCaps ? 'uppercase' : undefined,
        fontVariant: style.smallCaps ? 'small-caps' : undefined,
        writingMode: style.writingMode,
        whiteSpace: 'pre-wrap',
        lineHeight: style.lineHeight ?? 1.15,
        background: element.fill ? colorWithOpacity(element.fill, element.fillOpacity) : undefined,
        border: element.stroke
          ? `${element.strokeWidth ?? 1}px ${lineStyle(element.strokeDash)} ${colorWithOpacity(element.stroke, element.strokeOpacity)}`
          : undefined,
        borderRadius: radiusToPx(element),
        boxShadow: shadowToCss(element),
        boxSizing: 'border-box',
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        justifyContent: style.verticalAlign === 'bottom' ? 'flex-end' : style.verticalAlign === 'middle' ? 'center' : 'flex-start',
        paddingLeft: style.marginLeft ?? 0,
        paddingRight: style.marginRight ?? 0,
        paddingTop: style.marginTop ?? 0,
        paddingBottom: style.marginBottom ?? 0,
        letterSpacing: 0,
        wordBreak: 'break-word',
        overflowWrap: 'anywhere',
      }}
    >
      {element.paragraphs.map((paragraph, paragraphIndex) => (
        <div
          key={paragraphIndex}
          style={{
            textAlign: paragraph.style?.align ?? style.align ?? 'left',
            lineHeight: paragraph.style?.lineHeight ?? style.lineHeight ?? 1.2,
            marginTop: paragraph.style?.spaceBefore ?? 0,
            marginBottom: paragraph.style?.spaceAfter ?? 0,
            paddingLeft: `${(paragraph.style?.marginLeft ?? 0) + (paragraph.bullet && !paragraph.bullet.none ? 18 : 0)}px`,
            textIndent: paragraph.style?.textIndent ? `${paragraph.style.textIndent}px` : undefined,
            display: 'block',
            whiteSpace: 'inherit',
          }}
        >
          {paragraph.bullet && !paragraph.bullet.none ? (
            <span
              style={{
                display: 'inline-block',
                color: colorWithOpacity(paragraph.bullet.color ?? style.color, style.opacity),
                fontSize: paragraph.bullet.size ?? style.fontSize,
                marginRight: 6,
                width: 12,
                textAlign: 'center',
              }}
            >
              {paragraph.bullet.char ?? '\u2022'}
            </span>
          ) : null}
          {paragraph.runs.map((run, runIndex) => {
            const runStyle = run.style ?? {};
            return (
              <span
                key={runIndex}
                style={{
                  color: colorWithOpacity(runStyle.color ?? style.color ?? '#172033', runStyle.opacity ?? style.opacity),
                  fontFamily: runStyle.fontFamily ?? style.fontFamily,
                  fontSize: runStyle.fontSize ?? style.fontSize,
                  fontWeight: runStyle.bold || style.bold ? 600 : 400,
                  fontStyle: runStyle.italic || style.italic ? 'italic' : 'normal',
                  textDecoration: [
                    runStyle.underline || style.underline ? 'underline' : '',
                    runStyle.strike && runStyle.strike !== 'none' ? 'line-through' : '',
                  ]
                    .filter(Boolean)
                    .join(' ') || 'none',
                  textTransform: runStyle.allCaps || style.allCaps ? 'uppercase' : undefined,
                  fontVariant: runStyle.smallCaps || style.smallCaps ? 'small-caps' : undefined,
                  verticalAlign:
                    runStyle.baseline && runStyle.baseline > 0
                      ? 'super'
                      : runStyle.baseline && runStyle.baseline < 0
                        ? 'sub'
                        : undefined,
                  letterSpacing: 0,
                }}
              >
                {run.text}
              </span>
            );
          })}
        </div>
      ))}
    </div>
  );
}
