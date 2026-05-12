// PptxViewer 负责 PPTX 预览整体布局，组合左侧缩略图栏和右侧幻灯片视口。
import { memo } from 'react';
import type { PptxDocument } from '../../services/pptx/types';
import { OfficeEmpty } from '../office-viewer/OfficeEmpty';
import './index.less';
import { PptxSlideViewport } from './PptxSlideViewport';
import { PptxThumbnailPane } from './PptxThumbnailPane';

type PptxViewerProps = {
  document?: PptxDocument;
  activeIndex: number;
  zoom: number;
  onSelectSlide: (index: number) => void;
};

function PptxViewerComponent({ document, activeIndex, zoom, onSelectSlide }: PptxViewerProps) {
  if (!document?.slides.length) {
    return <OfficeEmpty kind="pptx" />;
  }

  const currentSlide = document.slides[activeIndex];

  return (
    <div className="oxv-pptx-viewer">
      <PptxThumbnailPane slides={document.slides} activeIndex={activeIndex} onSelectSlide={onSelectSlide} />
      <PptxSlideViewport slide={currentSlide} activeIndex={activeIndex} zoom={zoom} />
    </div>
  );
}

export const PptxViewer = memo(PptxViewerComponent);
