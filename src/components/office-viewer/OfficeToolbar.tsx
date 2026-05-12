import { Button, Select, Space, Tooltip, Typography, Upload } from 'antd';
import { memo } from 'react';
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
import type { PreviewKind } from '../../services/officePreview';
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
  if (kind === 'docx') return <FileWordOutlined />;
  return <FilePptOutlined />;
}

function OfficeToolbarComponent({
  fileName,
  previewKind,
  uploadAccept = '.pptx,.xlsx,.docx',
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
  return (
    <div
      style={{
        height: 56,
        background: '#fff',
        borderBottom: '1px solid #dde3ec',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: 16,
        padding: '0 16px',
        position: 'sticky',
        top: 0,
        zIndex: 20,
      }}
    >
      <Typography.Text strong ellipsis style={{ maxWidth: 360 }}>
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
        <Select
          value={zoom}
          style={{ width: 104 }}
          onChange={onZoomChange}
          options={OFFICE_ZOOM_LEVELS.map((value) => ({ value, label: `${value}%` }))}
        />
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

