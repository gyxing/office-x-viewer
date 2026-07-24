// OfficeChartView 将解析后的 Office 图表模型渲染为 ECharts 图表。
import { Empty, Spin } from 'antd';
import type { CSSProperties } from 'react';
import React, { memo, useEffect, useMemo, useRef, useState } from 'react';
import type { OfficeChartModel } from '../../shared/ooxml/charts';
import { buildOfficeChartOption } from '../../shared/ooxml/charts';
import './index.less';

type OfficeChartViewProps = {
  chart: OfficeChartModel;
  width: number;
  height: number;
  zoom?: number;
};

// 地图 GeoJSON 注册到 ECharts 后是全局状态，同一个 mapName 不需要重复下载和注册。
const registeredMaps = new Set<string>();

function OfficeChartViewComponent({
  chart,
  width,
  height,
  zoom = 100,
}: OfficeChartViewProps) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const chartRef = useRef<import('echarts').EChartsType | null>(null);
  const echartsRef = useRef<typeof import('echarts') | null>(null);
  const [ready, setReady] = useState(false);
  const [visible, setVisible] = useState(false);
  const [mapFailed, setMapFailed] = useState(false);
  const [renderFailed, setRenderFailed] = useState(false);
  const [generatedSnapshotSrc, setGeneratedSnapshotSrc] = useState<
    string | undefined
  >();
  const snapshotCapturedRef = useRef(false);
  const displayWidth = width * (zoom / 100);
  const displayHeight = height * (zoom / 100);

  const outerStyle = useMemo<CSSProperties>(
    () => ({
      width: displayWidth,
      height: displayHeight,
    }),
    [displayHeight, displayWidth],
  );

  useEffect(() => {
    setGeneratedSnapshotSrc(undefined);
    snapshotCapturedRef.current = false;
  }, [chart]);

  useEffect(() => {
    setMapFailed(false);
    setRenderFailed(false);
    setReady(false);
    let disposed = false;
    const requestController =
      typeof AbortController === 'undefined'
        ? undefined
        : new AbortController();
    let resizeObserver: ResizeObserver | undefined;
    let intersectionObserver: IntersectionObserver | undefined;

    // 图表可能出现在缩略图、表格浮层或文档深处，进入视口后再加载 ECharts，减少首屏成本。
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
      if (
        !visible ||
        !hostRef.current ||
        chartRef.current ||
        (chart.renderMode === 'snapshot' && chart.snapshotSrc)
      ) {
        return;
      }

      try {
        const echarts = echartsRef.current ?? (await import('echarts'));
        echartsRef.current = echarts;
        if (chart.type === 'map') {
          const mapName = chart.mapName ?? 'china';
          if (!registeredMaps.has(mapName)) {
            if (!chart.mapGeoJsonUrl) {
              setMapFailed(true);
              return;
            }

            try {
              // 地图图表需要额外 GeoJSON；失败时优先回退到 Office/WPS 里携带的快照图。
              const response = await fetch(chart.mapGeoJsonUrl, {
                signal: requestController?.signal,
              });
              if (!response.ok)
                throw new Error(`Map data request failed: ${response.status}`);
              const geoJson = await response.json();
              if (disposed) return;
              echarts.registerMap(mapName, geoJson);
              registeredMaps.add(mapName);
              setMapFailed(false);
            } catch {
              if (!disposed && !requestController?.signal.aborted) {
                setMapFailed(true);
              }
              return;
            }
          }
        }
        if (disposed || !hostRef.current) return;
        const instance = echarts.init(hostRef.current, undefined, {
          renderer: 'canvas',
        });
        chartRef.current = instance;
        if (chart.renderMode === 'snapshot') {
          const captureSnapshot = () => {
            if (snapshotCapturedRef.current || disposed) return;
            snapshotCapturedRef.current = true;
            try {
              const snapshot = instance.getDataURL({
                type: 'png',
                pixelRatio: 1,
                backgroundColor: '#fff',
              });
              setGeneratedSnapshotSrc(snapshot);
              instance.dispose();
              chartRef.current = null;
            } catch {
              if (!chart.snapshotSrc) setRenderFailed(true);
            }
          };
          instance.on('finished', captureSnapshot);
        }
        instance.setOption(buildOfficeChartOption(chart), {
          notMerge: true,
          lazyUpdate: chart.renderMode !== 'snapshot',
        });
        setReady(true);

        // 外层会随 zoom 和文档布局变化；旧浏览器没有 ResizeObserver 时保留初始化尺寸。
        if (typeof ResizeObserver !== 'undefined') {
          resizeObserver = new ResizeObserver(() => {
            instance.resize();
          });
          resizeObserver.observe(hostRef.current);
        }
      } catch {
        if (disposed || requestController?.signal.aborted) return;

        chartRef.current?.dispose();
        chartRef.current = null;
        setReady(false);
        setRenderFailed(true);
      }
    }

    void mountChart();

    return () => {
      disposed = true;
      requestController?.abort();
      resizeObserver?.disconnect();
      intersectionObserver?.disconnect();
      chartRef.current?.dispose();
      chartRef.current = null;
      setReady(false);
    };
  }, [chart, visible]);

  if (!width || !height) {
    return <Empty description="图表尺寸无效" />;
  }

  const staticSnapshotSrc =
    chart.renderMode === 'snapshot'
      ? chart.snapshotSrc ?? generatedSnapshotSrc
      : undefined;
  if (staticSnapshotSrc) {
    return (
      <div
        className="oxv-office-chart oxv-office-chart--static"
        style={outerStyle}
      >
        <img
          className="oxv-office-chart__snapshot"
          src={staticSnapshotSrc}
          alt={chart.title ?? chart.degradedFrom ?? '静态图表'}
        />
      </div>
    );
  }

  if (renderFailed || (chart.type === 'map' && mapFailed)) {
    if (!chart.snapshotSrc) {
      return (
        <Empty
          description={renderFailed ? '图表渲染失败' : '地图数据加载失败'}
        />
      );
    }

    return (
      <div className="oxv-office-chart" style={outerStyle}>
        <img
          className="oxv-office-chart__snapshot"
          src={chart.snapshotSrc}
          alt={chart.title ?? chart.mapRegion ?? chart.degradedFrom ?? ''}
        />
      </div>
    );
  }

  return (
    <div className="oxv-office-chart" style={outerStyle}>
      <div ref={hostRef} className="oxv-office-chart__host" />
      {!ready ? (
        <div className="oxv-office-chart__loading">
          <Spin />
        </div>
      ) : null}
    </div>
  );
}

export const OfficeChartView = memo(OfficeChartViewComponent);
