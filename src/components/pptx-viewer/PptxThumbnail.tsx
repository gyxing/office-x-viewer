import type { SlideModel } from '../../services/pptx/types';
import { PptxSlide } from './PptxSlide';

type PptxThumbnailProps = {
  slide: SlideModel;
  active: boolean;
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

export function PptxThumbnail({ slide, active }: PptxThumbnailProps) {
  const ratio = slide.width / slide.height;

  return (
    <div
      style={{
        border: active ? '1px solid #2f6fed' : '1px solid #dde3ec',
        borderRadius: 6,
        padding: 8,
        background: '#fff',
        boxShadow: active ? '0 0 0 2px rgba(47, 111, 237, 0.12)' : 'none',
      }}
    >
      <div
        style={{
          aspectRatio: `${ratio}`,
          position: 'relative',
          overflow: 'hidden',
          borderRadius: 4,
          background: colorWithOpacity(slide.background?.fill ?? '#f8fafc', slide.background?.fillOpacity),
        }}
      >
        {slide.background?.imageRef ? (
          <div
            style={{
              position: 'absolute',
              inset: 0,
              backgroundImage: `url(${slide.background.imageRef})`,
              backgroundSize: 'cover',
              backgroundPosition: 'center',
            }}
          />
        ) : null}
        <div
          style={{
            transform: 'scale(0.18)',
            transformOrigin: 'top left',
            width: `${100 / 0.18}%`,
            height: `${100 / 0.18}%`,
            pointerEvents: 'none',
            filter: 'saturate(0.98) contrast(0.98)',
          }}
        >
          <PptxSlide slide={slide} zoom={100} />
        </div>
      </div>
      <div style={{ marginTop: 8, fontSize: 12, color: '#667085' }}>第 {slide.index} 页</div>
    </div>
  );
}
