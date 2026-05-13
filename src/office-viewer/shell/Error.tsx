// OfficeError 展示文件解析或预览过程中的错误信息。
import { Alert } from 'antd';
import { memo } from 'react';

type OfficeErrorProps = {
  message: string;
};

function OfficeErrorComponent({ message }: OfficeErrorProps) {
  return (
    <div className="oxv-office-error">
      <Alert type="error" showIcon message="预览失败" description={message} />
    </div>
  );
}

export const OfficeError = memo(OfficeErrorComponent);
