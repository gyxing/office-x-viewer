import { memo, useEffect, useRef } from 'react';
import type { SlideModel } from '../../services/pptx/types';
import { PptxSlide } from './PptxSlide';

type PptxSlideViewportProps = {
  slide?: SlideModel;
  activeIndex: number;
  zoom: number;
};

function PptxSlideViewportComponent({ slide, activeIndex, zoom }: PptxSlideViewportProps) {
  const viewportRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    viewportRef.current?.scrollTo({ left: 0, top: 0 });
  }, [activeIndex, zoom]);

  return (
    <section ref={viewportRef} className="oxv-pptx-viewer__viewport">
      <div className="oxv-pptx-viewer__slide-wrap">
        {slide ? <PptxSlide slide={slide} zoom={zoom} renderKey={`slide-${slide.id}`} /> : null}
      </div>
    </section>
  );
}

export const PptxSlideViewport = memo(PptxSlideViewportComponent);
