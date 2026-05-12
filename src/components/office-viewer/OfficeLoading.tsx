import { Spin } from 'antd';
import { memo } from 'react';

type OfficeLoadingProps = {
  tip?: string;
};

function OfficeLoadingComponent({ tip = '正在解析文件' }: OfficeLoadingProps) {
  return (
    <div style={{ height: 'calc(100vh - 56px)', display: 'grid', placeItems: 'center' }}>
      <Spin size="large" tip={tip} />
    </div>
  );
}

export const OfficeLoading = memo(OfficeLoadingComponent);
