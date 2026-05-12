import { memo } from 'react';
import type { SlideElement, SlideModel } from '../../services/pptx/types';
import { colorWithOpacity } from './renderers/paint';
import { ShapeRenderer } from './renderers/ShapeRenderer';
import { TextRenderer } from './renderers/TextRenderer';
import { ImageRenderer } from './renderers/ImageRenderer';
import { TableRenderer } from './renderers/TableRenderer';
import { OfficeChartView } from '../office-chart/OfficeChartView';
import { UnsupportedRenderer } from './renderers/UnsupportedRenderer';

type PptxSlideProps = {
  slide: SlideModel;
  zoom: number;
  renderKey?: string;
};

const ChartFrame = memo(function ChartFrame({
  element,
  zoom,
}: {
  element: Extract<SlideElement, { type: 'chart' }>;
  zoom: number;
}) {
  return (
    <div
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
      }}
    >
      <OfficeChartView chart={element.chart} width={element.width} height={element.height} zoom={zoom} />
    </div>
  );
});

function PptxSlideComponent({ slide, zoom, renderKey }: PptxSlideProps) {
  const scale = zoom / 100;
  const slideRenderKey = renderKey ?? `slide-${slide.id}`;

  return (
    <div
      style={{
        width: slide.width,
        height: slide.height,
        minWidth: slide.width,
        minHeight: slide.height,
        position: 'relative',
        overflow: 'hidden',
        boxShadow: '0 12px 32px rgba(15, 23, 42, 0.16)',
        transform: `scale(${scale})`,
        transformOrigin: 'top left',
        isolation: 'isolate',
      }}
    >
      <div
        style={{
          position: 'absolute',
          inset: 0,
          background: colorWithOpacity(slide.background?.fill ?? '#fff', slide.background?.fillOpacity),
          backgroundImage: slide.background?.imageRef ? `url(${slide.background.imageRef})` : undefined,
          backgroundSize: 'cover',
          backgroundPosition: 'center',
          zIndex: 0,
        }}
      />
      <div style={{ position: 'absolute', inset: 0, zIndex: 1 }}>
        {slide.elements.map((element) => {
          switch (element.type) {
            case 'text':
              return <TextRenderer key={element.id} element={element} renderKey={slideRenderKey} />;
            case 'shape':
              return <ShapeRenderer key={element.id} element={element} renderKey={slideRenderKey} />;
            case 'image':
              return <ImageRenderer key={element.id} element={element} />;
            case 'table':
              return <TableRenderer key={element.id} element={element} />;
            case 'chart':
              return <ChartFrame key={element.id} element={element} zoom={zoom} />;
            case 'unsupported':
              return <UnsupportedRenderer key={element.id} element={element} />;
            default:
              return null;
          }
        })}
      </div>
    </div>
  );
}

export const PptxSlide = memo(PptxSlideComponent);
