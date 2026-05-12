import { memo } from 'react';
import type { SlideModel } from '../../services/pptx/types';
import { colorWithOpacity } from './renderers/paint';
import { PptxSlide } from './PptxSlide';

type PptxThumbnailProps = {
  slide: SlideModel;
  active: boolean;
};

function PptxThumbnailComponent({ slide, active }: PptxThumbnailProps) {
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
          <PptxSlide slide={slide} zoom={100} renderKey={`thumb-${slide.id}`} />
        </div>
      </div>
      <div style={{ marginTop: 8, fontSize: 12, color: '#667085' }}>第 {slide.index} 页</div>
    </div>
  );
}

export const PptxThumbnail = memo(PptxThumbnailComponent);

