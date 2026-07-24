import type { OfficeChartType } from '../../../shared/ooxml/charts';
import type { AdaptedBiff8Chart, Biff8ChartModel } from './types';

const COMMON_TYPES = new Set([
  'column',
  'bar',
  'line',
  'pie',
  'doughnut',
  'area',
  'scatter',
  'bubble',
  'radar',
  'radarArea',
]);

function degradedType(sourceType: string): OfficeChartType {
  if (sourceType === 'surface') return 'area';
  if (sourceType === 'stock' || sourceType === 'unknown') return 'line';
  if (sourceType === 'radarArea') return 'radar';
  if (sourceType === 'ofPie') return 'pie';
  return COMMON_TYPES.has(sourceType)
    ? (sourceType as OfficeChartType)
    : 'line';
}

/** 将 BIFF8 图表模型映射到共享 Office/ECharts 图表模型。 */
export function adaptBiff8Chart(source: Biff8ChartModel): AdaptedBiff8Chart {
  const unsupportedType =
    !COMMON_TYPES.has(source.sourceType) ||
    source.groupTypes.some((type) => !COMMON_TYPES.has(type));
  const renderMode =
    source.is3d || source.hasSecondaryAxis || unsupportedType
      ? ('snapshot' as const)
      : ('interactive' as const);
  const degradedFrom =
    renderMode === 'snapshot'
      ? `${source.is3d ? '3D ' : ''}${
          source.hasSecondaryAxis ? '次坐标 ' : ''
        }${source.sourceType}`
      : undefined;
  const chartType = degradedType(source.sourceType);
  return {
    renderMode,
    degradedFrom,
    warnings: source.warnings,
    chart: {
      type: chartType,
      title: source.title,
      categories: source.categories,
      series: source.series.map((series) => ({
        name: series.name,
        type:
          series.type ??
          degradedType(
            source.groupTypes[series.groupIndex] ?? source.sourceType,
          ),
        values: series.values.map((value) => value ?? 0),
        xValues: series.categories.map((value, index) =>
          typeof value === 'number' ? value : index + 1,
        ),
        bubbleSizes: series.bubbleSizes.map((value) => value ?? 0),
        stacking: series.stacking,
        stackGroup: series.stacking
          ? `xls-stack-${series.groupIndex}`
          : undefined,
        gapWidth: source.gapWidth,
        overlap: source.overlap,
        color: series.color,
        marker: series.marker,
        lineWidth: series.lineWidth,
      })),
      showLegend: source.showLegend,
      legendPosition: source.legendPosition,
      gapWidth: source.gapWidth,
      overlap: source.overlap,
      holeSize: source.holeSize,
      radarStyle: source.sourceType === 'radarArea' ? 'filled' : undefined,
      sourceType: source.sourceType,
      renderMode,
      degradedFrom,
      snapshotSrc: source.previewImageSrc,
    },
  };
}
