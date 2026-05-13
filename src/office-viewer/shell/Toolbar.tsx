// OfficeToolbar 提供选择文件、翻页、缩放、全屏等 OfficeViewer 顶部通用操作。
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
import type { PreviewKind } from '../services/preview';
import {
  OFFICE_DEFAULT_ZOOM,
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_LEVELS,
} from './constants';

const OFFICE_FILE_ACCEPT = '.pptx,.xlsx,.docx,.doc';

type OfficeToolbarProps = {
  fileName: string;
  previewKind: PreviewKind;
  zoom: number;
  hasDocument: boolean;
  canGoPreviousSlide: boolean;
  canGoNextSlide: boolean;
  onSelectFile: (file: File) => void;
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
  zoom,
  hasDocument,
  canGoPreviousSlide,
  canGoNextSlide,
  onSelectFile,
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
          accept={OFFICE_FILE_ACCEPT}
          showUploadList={false}
          beforeUpload={(file) => {
            void onSelectFile(file);
            return false;
          }}
        >
          <Button icon={getPreviewIcon(previewKind)}>选择文件</Button>
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

