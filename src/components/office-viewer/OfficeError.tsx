import { Alert } from 'antd';

type OfficeErrorProps = {
  message: string;
};

export function OfficeError({ message }: OfficeErrorProps) {
  return (
    <div style={{ padding: 24 }}>
      <Alert type="error" showIcon message="预览失败" description={message} />
    </div>
  );
}

