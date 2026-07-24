import React, { memo } from 'react';
import type { ParseProgress } from '../services/parsing';
import { OfficeNotice } from './Notice';

type OfficeParseStatusProps = {
  progress?: ParseProgress;
  warning?: string;
};

function normalizePercent(progress: ParseProgress | undefined) {
  if (progress?.percent !== undefined) {
    return Math.max(0, Math.min(100, progress.percent * 100));
  }
  if (progress?.total && progress.completed !== undefined) {
    return Math.max(
      0,
      Math.min(100, (progress.completed / progress.total) * 100),
    );
  }
  return undefined;
}

/** OfficeParseStatus 在预览内容上方展示非阻塞解析进度或不完整警告。 */
function OfficeParseStatusComponent({
  progress,
  warning,
}: OfficeParseStatusProps) {
  if (warning) {
    return (
      <div className="oxv-office-parse-status" role="alert">
        <OfficeNotice
          type="warning"
          title="文档解析未完成"
          description={`当前仅展示已成功解析的部分内容。失败原因：${warning}`}
        />
      </div>
    );
  }
  if (!progress) return null;

  const percent = normalizePercent(progress);
  const barClassName = [
    'oxv-office-parse-status__bar',
    percent === undefined
      ? 'oxv-office-parse-status__bar--indeterminate'
      : '',
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <div
      className="oxv-office-parse-status"
      role="status"
      aria-live="polite"
    >
      <div className="oxv-office-parse-status__label">
        {progress.message}
      </div>
      <div
        className="oxv-office-parse-status__track"
        role="progressbar"
        aria-valuemin={0}
        aria-valuemax={100}
        aria-valuenow={percent === undefined ? undefined : Math.round(percent)}
      >
        <div
          className={barClassName}
          style={
            percent === undefined ? undefined : { width: `${percent}%` }
          }
        />
      </div>
    </div>
  );
}

export const OfficeParseStatus = memo(OfficeParseStatusComponent);
