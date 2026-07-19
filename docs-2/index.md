# office-x-viewer

支持 `antd >=4.24.0 <7.0.0`。antd 4/5 可搭配 React 16.9+；antd 6 需要 React 18+。

使用 antd v4 时，请在宿主入口加载：

```tsx | pure
import 'antd/dist/antd.css';
```

```tsx
import React from 'react';
import { OfficeViewer } from 'office-x-viewer';

export default () => {
  return <OfficeViewer />;
};
```
