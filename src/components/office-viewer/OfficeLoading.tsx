import { Spin } from 'antd';

type OfficeLoadingProps = {
  tip?: string;
};

export function OfficeLoading({ tip = '正在解析文件' }: OfficeLoadingProps) {
  return (
    <div style={{ height: 'calc(100vh - 56px)', display: 'grid', placeItems: 'center' }}>
      <Spin size="large" tip={tip} />
    </div>
  );
}

