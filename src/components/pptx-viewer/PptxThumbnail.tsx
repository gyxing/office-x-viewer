import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { SlideModel } from '../../services/pptx/types';
import { colorWithOpacity } from './renderers/paint';
import { PptxSlide } from './PptxSlide';

type PptxThumbnailProps = {
  slide: SlideModel;
  active: boolean;
};

function PptxThumbnailComponent({ slide, active }: PptxThumbnailProps) {
  const canvasStyle = useMemo<CSSProperties>(
    () => ({
      aspectRatio: `${slide.width / slide.height}`,
      background: colorWithOpacity(slide.background?.fill ?? '#f8fafc', slide.background?.fillOpacity),
    }),
    [slide.background?.fill, slide.background?.fillOpacity, slide.height, slide.width],
  );
  const backgroundStyle = useMemo<CSSProperties>(
    () => ({
      backgroundImage: slide.background?.imageRef ? `url(${slide.background.imageRef})` : undefined,
    }),
    [slide.background?.imageRef],
  );

  return (
    <div className={['oxv-pptx-thumbnail', active ? 'oxv-pptx-thumbnail--active' : ''].filter(Boolean).join(' ')}>
      <div className="oxv-pptx-thumbnail__canvas" style={canvasStyle}>
        {slide.background?.imageRef ? <div className="oxv-pptx-thumbnail__background" style={backgroundStyle} /> : null}
        <div className="oxv-pptx-thumbnail__content">
          <PptxSlide slide={slide} zoom={100} renderKey={`thumb-${slide.id}`} />
        </div>
      </div>
      <div className="oxv-pptx-thumbnail__label">第 {slide.index} 页</div>
    </div>
  );
}

export const PptxThumbnail = memo(PptxThumbnailComponent);
