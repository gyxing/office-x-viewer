import { Alert } from 'antd';
import { memo } from 'react';

type OfficeErrorProps = {
  message: string;
};

function OfficeErrorComponent({ message }: OfficeErrorProps) {
  return (
    <div style={{ padding: 24 }}>
      <Alert type="error" showIcon message="预览失败" description={message} />
    </div>
  );
}

export const OfficeError = memo(OfficeErrorComponent);
