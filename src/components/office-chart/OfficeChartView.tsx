import { Empty, Spin } from 'antd';
import type { CSSProperties } from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { OfficeChartModel } from '../../services/office/charts';
import { buildOfficeChartOption } from '../../services/office/charts';

type OfficeChartViewProps = {
  chart: OfficeChartModel;
  width: number;
  height: number;
  zoom?: number;
};

export function OfficeChartView({ chart, width, height, zoom = 100 }: OfficeChartViewProps) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const chartRef = useRef<import('echarts').EChartsType | null>(null);
  const echartsRef = useRef<typeof import('echarts') | null>(null);
  const [ready, setReady] = useState(false);
  const [visible, setVisible] = useState(false);
  const displayWidth = width * (zoom / 100);
  const displayHeight = height * (zoom / 100);

  const style = useMemo<CSSProperties>(
    () => ({
      width: '100%',
      height: '100%',
      minWidth: 0,
      minHeight: 0,
    }),
    [],
  );

  const outerStyle = useMemo<CSSProperties>(
    () => ({
      position: 'relative',
      width: displayWidth,
      height: displayHeight,
      overflow: 'hidden',
    }),
    [displayHeight, displayWidth],
  );

  useEffect(() => {
    let disposed = false;
    let resizeObserver: ResizeObserver | undefined;
    let intersectionObserver: IntersectionObserver | undefined;

    if (hostRef.current && typeof IntersectionObserver !== 'undefined') {
      intersectionObserver = new IntersectionObserver(
        (entries) => {
          setVisible(entries.some((entry) => entry.isIntersecting));
        },
        { threshold: 0.01 },
      );
      intersectionObserver.observe(hostRef.current);
    } else {
      setVisible(true);
    }

    async function mountChart() {
      if (!visible || !hostRef.current || chartRef.current) return;
      const echarts = echartsRef.current ?? (await import('echarts'));
      echartsRef.current = echarts;
      if (disposed || !hostRef.current) return;
      const instance = echarts.init(hostRef.current, undefined, { renderer: 'canvas' });
      chartRef.current = instance;
      instance.setOption(buildOfficeChartOption(chart), { notMerge: true, lazyUpdate: true });
      setReady(true);

      resizeObserver = new ResizeObserver(() => {
        instance.resize();
      });
      resizeObserver.observe(hostRef.current);
    }

    void mountChart();

    return () => {
      disposed = true;
      resizeObserver?.disconnect();
      intersectionObserver?.disconnect();
      chartRef.current?.dispose();
      chartRef.current = null;
      setReady(false);
    };
  }, [chart, visible]);

  useEffect(() => {
    if (!chartRef.current) return;
    chartRef.current.setOption(buildOfficeChartOption(chart), { notMerge: true, lazyUpdate: true });
    chartRef.current.resize();
  }, [chart]);

  if (!width || !height) {
    return <Empty description="图表尺寸无效" />;
  }

  return (
    <div style={outerStyle}>
      <div ref={hostRef} style={style} />
      {!ready ? (
        <div
          style={{
            position: 'absolute',
            inset: 0,
            display: 'grid',
            placeItems: 'center',
            background: 'rgba(255,255,255,0.75)',
          }}
        >
          <Spin />
        </div>
      ) : null}
    </div>
  );
}
