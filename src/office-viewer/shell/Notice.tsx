// OfficeNotice 展示跨 antd 版本稳定的错误或警告状态。
import React, { memo } from 'react';

type OfficeNoticeProps = {
  type: 'error' | 'warning';
  title: string;
  description: string;
};

function OfficeNoticeComponent({
  type,
  title,
  description,
}: OfficeNoticeProps) {
  return (
    <div
      className={`oxv-office-notice oxv-office-notice--${type}`}
      role="alert"
    >
      <span className="oxv-office-notice__icon" aria-hidden="true">
        !
      </span>
      <div className="oxv-office-notice__content">
        <div className="oxv-office-notice__title">{title}</div>
        <div className="oxv-office-notice__description">{description}</div>
      </div>
    </div>
  );
}

export const OfficeNotice = memo(OfficeNoticeComponent);
