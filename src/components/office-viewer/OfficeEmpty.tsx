import { Empty } from 'antd';
import type { PreviewKind } from '../../services/officePreview';
import { OFFICE_EMPTY_DESCRIPTIONS } from './shared/constants';

type OfficeEmptyProps = {
  kind: PreviewKind;
};

export function OfficeEmpty({ kind }: OfficeEmptyProps) {
  return <Empty description={OFFICE_EMPTY_DESCRIPTIONS[kind]} />;
}

