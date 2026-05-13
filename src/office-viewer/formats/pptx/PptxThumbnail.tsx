// PptxThumbnail 复用单页幻灯片渲染能力，生成缩略图预览。
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
          {/* 缩略图复用完整 Slide 渲染，保证背景、图形、表格和图表与主画布一致。 */}
          <PptxSlide slide={slide} zoom={100} renderKey={`thumb-${slide.id}`} />
        </div>
      </div>
      <div className="oxv-pptx-thumbnail__label">第 {slide.index} 页</div>
    </div>
  );
}

export const PptxThumbnail = memo(PptxThumbnailComponent);
