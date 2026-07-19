// OfficeError 展示文件解析或预览过程中的错误信息。
import React, { memo } from 'react';
import { OfficeNotice } from './Notice';

type OfficeErrorProps = {
  message: string;
};

function OfficeErrorComponent({ message }: OfficeErrorProps) {
  return (
    <div className="oxv-office-error">
      <OfficeNotice type="error" title="预览失败" description={message} />
    </div>
  );
}

export const OfficeError = memo(OfficeErrorComponent);
