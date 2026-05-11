import type { SlideModel } from '../../services/pptx/types';
import { ShapeRenderer } from './renderers/ShapeRenderer';
import { TextRenderer } from './renderers/TextRenderer';
import { ImageRenderer } from './renderers/ImageRenderer';
import { TableRenderer } from './renderers/TableRenderer';
import { OfficeChartView } from '../office-chart/OfficeChartView';
import { UnsupportedRenderer } from './renderers/UnsupportedRenderer';

type PptxSlideProps = {
  slide: SlideModel;
  zoom: number;
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

function ChartFrame({
  element,
  zoom,
}: {
  element: Extract<import('../../services/pptx/types').SlideElement, { type: 'chart' }>;
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
}

export function PptxSlide({ slide, zoom }: PptxSlideProps) {
  const scale = zoom / 100;

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
              return <TextRenderer key={element.id} element={element} />;
            case 'shape':
              return <ShapeRenderer key={element.id} element={element} />;
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
