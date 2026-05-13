// OfficeToolbar 提供上传、翻页、缩放、全屏等 OfficeViewer 顶部通用操作。
import {
  FileExcelOutlined,
  FilePptOutlined,
  FileWordOutlined,
  FullscreenOutlined,
  LeftOutlined,
  RightOutlined,
  ZoomInOutlined,
  ZoomOutOutlined,
} from '@ant-design/icons';
import { Button, Select, Space, Tooltip, Typography, Upload } from 'antd';
import { memo, useMemo } from 'react';
import type { PreviewKind } from '../../services/office/preview';
import {
  OFFICE_DEFAULT_ZOOM,
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_LEVELS,
} from './shared/constants';

type OfficeToolbarProps = {
  fileName: string;
  previewKind: PreviewKind;
  uploadAccept?: string;
  uploadLabel?: string;
  zoom: number;
  hasDocument: boolean;
  canGoPreviousSlide: boolean;
  canGoNextSlide: boolean;
  onUpload: (file: File) => void;
  onPreviousSlide: () => void;
  onNextSlide: () => void;
  onZoomOut: () => void;
  onZoomIn: () => void;
  onZoomChange: (zoom: number) => void;
  onResetZoom: () => void;
  onFullscreen?: () => void;
};

function getPreviewIcon(kind: PreviewKind) {
  if (kind === 'xlsx') return <FileExcelOutlined />;
  if (kind === 'docx' || kind === 'doc') return <FileWordOutlined />;
  return <FilePptOutlined />;
}

function OfficeToolbarComponent({
  fileName,
  previewKind,
  uploadAccept = '.pptx,.xlsx,.docx,.doc',
  uploadLabel = '上传文件',
  zoom,
  hasDocument,
  canGoPreviousSlide,
  canGoNextSlide,
  onUpload,
  onPreviousSlide,
  onNextSlide,
  onZoomOut,
  onZoomIn,
  onZoomChange,
  onResetZoom,
  onFullscreen,
}: OfficeToolbarProps) {
  const zoomOptions = useMemo(() => OFFICE_ZOOM_LEVELS.map((value) => ({ value, label: `${value}%` })), []);

  return (
    <div className="oxv-office-toolbar">
      <Typography.Text strong ellipsis className="oxv-office-toolbar__filename">
        {fileName}
      </Typography.Text>
      <Space size={8} wrap>
        <Upload
          accept={uploadAccept}
          showUploadList={false}
          beforeUpload={(file) => {
            void onUpload(file);
            return false;
          }}
        >
          <Button icon={getPreviewIcon(previewKind)}>{uploadLabel}</Button>
        </Upload>
        <Tooltip title="上一页">
          <Button
            icon={<LeftOutlined />}
            disabled={previewKind !== 'pptx' || !hasDocument || !canGoPreviousSlide}
            onClick={onPreviousSlide}
          />
        </Tooltip>
        <Tooltip title="下一页">
          <Button
            icon={<RightOutlined />}
            disabled={previewKind !== 'pptx' || !hasDocument || !canGoNextSlide}
            onClick={onNextSlide}
          />
        </Tooltip>
        <Select value={zoom} className="oxv-office-toolbar__zoom" onChange={onZoomChange} options={zoomOptions} />
        <Tooltip title="缩小">
          <Button
            icon={<ZoomOutOutlined />}
            disabled={!hasDocument || zoom <= OFFICE_MIN_ZOOM}
            onClick={onZoomOut}
          />
        </Tooltip>
        <Tooltip title="放大">
          <Button
            icon={<ZoomInOutlined />}
            disabled={!hasDocument || zoom >= OFFICE_MAX_ZOOM}
            onClick={onZoomIn}
          />
        </Tooltip>
        <Button disabled={!hasDocument} onClick={onResetZoom}>
          {OFFICE_DEFAULT_ZOOM}%
        </Button>
        <Button icon={<FullscreenOutlined />} disabled={!hasDocument || !onFullscreen} onClick={onFullscreen}>
          全屏
        </Button>
      </Space>
    </div>
  );
}

export const OfficeToolbar = memo(OfficeToolbarComponent);
