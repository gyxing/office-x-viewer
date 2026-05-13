// PptxThumbnailPane 渲染幻灯片缩略图列表，并负责切换当前页。
import { memo, useCallback } from 'react';
import type { SlideModel } from '../../services/pptx/types';
import { PptxThumbnail } from './PptxThumbnail';

type PptxThumbnailPaneProps = {
  slides: SlideModel[];
  activeIndex: number;
  onSelectSlide: (index: number) => void;
};

function PptxThumbnailPaneComponent({ slides, activeIndex, onSelectSlide }: PptxThumbnailPaneProps) {
  const handleSelect = useCallback(
    (index: number) => {
      onSelectSlide(index);
    },
    [onSelectSlide],
  );

  return (
    <aside className="oxv-pptx-viewer__sidebar">
      <div className="oxv-pptx-viewer__sidebar-header">
        <div className="oxv-pptx-viewer__slide-count">共 {slides.length} 页</div>
      </div>
      <div className="oxv-pptx-viewer__thumbnail-list">
        {slides.map((slide, index) => (
          <button
            key={slide.id}
            type="button"
            className="oxv-pptx-viewer__thumbnail-button"
            onClick={() => handleSelect(index)}
          >
            <PptxThumbnail slide={slide} active={index === activeIndex} />
          </button>
        ))}
      </div>
    </aside>
  );
}

export const PptxThumbnailPane = memo(PptxThumbnailPaneComponent);
