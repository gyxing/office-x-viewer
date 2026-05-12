import { Empty } from 'antd';
import { memo } from 'react';
import type { PreviewKind } from '../../services/officePreview';
import { OFFICE_EMPTY_DESCRIPTIONS } from './shared/constants';

type OfficeEmptyProps = {
  kind: PreviewKind;
};

function OfficeEmptyComponent({ kind }: OfficeEmptyProps) {
  return <Empty description={OFFICE_EMPTY_DESCRIPTIONS[kind]} />;
}

export const OfficeEmpty = memo(OfficeEmptyComponent);
