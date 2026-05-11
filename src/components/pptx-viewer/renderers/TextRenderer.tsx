import { useId } from 'react';
import type { TextElement } from '../../../services/pptx/types';
import { colorWithOpacity, gradientToSvgEndpoints, isGradientPaint, paintToCss } from './paint';

type TextRendererProps = {
  element: TextElement;
};

function shadowToCss(element: TextElement) {
  if (!element.shadow) return undefined;
  return `${element.shadow.offsetX ?? 0}px ${element.shadow.offsetY ?? 0}px ${Math.max(0, element.shadow.blur ?? 0)}px ${colorWithOpacity(element.shadow.color ?? '#000000', element.shadow.opacity ?? 0.18)}`;
}

function radiusToPx(element: TextElement) {
  if (element.shape === 'ellipse') return '50%';
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
  const instanceId = useId().replace(/[^a-zA-Z0-9_-]/g, '-');
  const style = element.boxStyle ?? {};
  const isVectorShape = Boolean(element.path);
  const fillPaint = element.fill;
  const isGradientFill = isGradientPaint(fillPaint);
  const gradientId = isGradientFill ? `${instanceId}-${element.id}-fill-gradient` : undefined;
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
        background: isVectorShape ? undefined : paintToCss(element.fill, element.fillOpacity),
        border: !isVectorShape && element.stroke
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
      {isVectorShape ? (
        <svg
          viewBox={element.viewBox ?? `0 0 ${Math.max(1, element.width)} ${Math.max(1, element.height)}`}
          preserveAspectRatio="none"
          style={{
            position: 'absolute',
            inset: 0,
            width: '100%',
            height: '100%',
            overflow: 'visible',
            zIndex: 0,
            pointerEvents: 'none',
          }}
        >
          {isGradientFill ? (
            <defs>
              <linearGradient id={gradientId} {...gradientToSvgEndpoints(fillPaint.angle)} gradientUnits="objectBoundingBox">
                {fillPaint.stops.map((stop, index) => (
                  <stop key={index} offset={`${stop.offset * 100}%`} stopColor={stop.color} />
                ))}
              </linearGradient>
            </defs>
          ) : null}
          <path
            d={element.path ?? ''}
            fill={isGradientFill ? `url(#${gradientId})` : element.fill ? colorWithOpacity(element.fill as string, element.fillOpacity) ?? 'none' : 'none'}
            fillOpacity={isGradientFill ? undefined : element.fillOpacity}
            stroke={element.stroke ? colorWithOpacity(element.stroke, element.strokeOpacity) ?? element.stroke : 'none'}
            strokeOpacity={element.strokeOpacity}
            strokeWidth={element.strokeWidth ?? 1}
            strokeDasharray={element.strokeDash && element.strokeDash !== 'solid' ? element.strokeDash : undefined}
            vectorEffect="non-scaling-stroke"
          />
        </svg>
      ) : null}
      {element.paragraphs.map((paragraph, paragraphIndex) => (
        <div
          key={paragraphIndex}
          style={{
            position: 'relative',
            zIndex: 1,
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
