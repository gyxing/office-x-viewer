# Office Viewer

`office-x-viewer` 是一个纯前端的 Microsoft Office 文档预览组件。文件下载、解析和渲染均在浏览器中完成，无需配套的文档转换服务。

组件基于 React、Ant Design、JSZip 和 ECharts，支持 PPTX、XLSX、DOCX，以及 DOC/WPS 的降级预览。

## 主要特性

- 统一的文件选择、加载、错误、空状态、缩放和全屏交互
- 支持本地 `File`、远程 URL、`Blob`、`Response` 和异步文件来源
- 切换 `uri` 时自动取消旧下载，并忽略已经过期的解析结果
- 兼容 antd v4、v5、v6，继承宿主项目的 `ConfigProvider`
- 功能性图标由组件内置 SVG 提供，不依赖 `@ant-design/icons`
- OOXML 文件完全在浏览器内解包和解析，ZIP 条目读取并发数固定为 4
- 图表按需加载；渲染或地图数据加载失败时优先回退到文档快照

## 安装

```bash
yarn add office-x-viewer antd react react-dom
```

`react`、`react-dom` 和 `antd` 由宿主项目提供。`echarts` 与 `jszip` 是组件自身的运行时依赖。

宿主构建工具需要支持 `.less` 文件，因为组件样式随模块一起导入。

## 版本兼容

| antd 版本           | React / ReactDOM | 支持状态 | 说明                              |
| ------------------- | ---------------- | -------- | --------------------------------- |
| `4.24.x`            | `>=16.9.0`       | 支持     | 宿主入口需要加载 antd v4 全局样式 |
| `5.x`               | `>=16.9.0`       | 支持     | 使用 antd v5 的样式机制           |
| `6.x`               | `>=18.0.0`       | 支持     | React 版本要求来自 antd v6        |
| `6.x` + React 16/17 | -                | 不支持   | 不满足 antd v6 自身要求           |

当前 peerDependencies 范围：

```text
antd: >=4.24.0 <7.0.0
react: >=16.9.0
react-dom: >=16.9.0
```

React 的 peer 下限用于兼容 antd v4/v5 宿主项目，并不代表 antd v6 可以运行在 React 16/17 上。

使用 antd v4 时，在宿主应用入口加载全局样式：

```tsx
import 'antd/dist/antd.css';
```

antd v5、v6 不需要导入上述文件。组件不会创建额外的根级 `ConfigProvider`，主题、语言和组件前缀由宿主配置决定。

## 基本用法

```tsx
import { OfficeViewer } from 'office-x-viewer';

export default function OfficePreview() {
  return <OfficeViewer />;
}
```

传入本地文件或远程地址：

```tsx
import { OfficeViewer } from 'office-x-viewer';

export default function OfficePreview({ file }: { file: File }) {
  return (
    <>
      <OfficeViewer uri={file} />
      <OfficeViewer uri="https://example.com/files/demo.pptx" />
    </>
  );
}
```

使用异步文件来源和回调：

```tsx
import { OfficeViewer } from 'office-x-viewer';

export default function OfficePreview() {
  return (
    <OfficeViewer
      uri={async () => fetch('/files/demo.xlsx')}
      onFileParsed={(parsed, file) => {
        console.info('解析完成', parsed.kind, file.name);
      }}
      onError={(error, file) => {
        console.error('预览失败', file?.name, error);
      }}
    />
  );
}
```

## `uri` 文件来源

`uri` 支持以下形式：

```ts
type OfficeViewerUri =
  | File
  | string
  | (() => Promise<File | Blob | string | Response>);
```

使用远程文件时需要注意：

- 跨域地址必须允许浏览器通过 CORS 访问。
- URL 最好包含受支持的文件扩展名。
- 无扩展名地址需要通过 `Content-Disposition` 文件名或响应 `Content-Type` 识别格式。
- 带有不受支持扩展名的 URL 会在下载前被拒绝，即使响应内容实际是 Office 文件。
- `uri` 变化时，旧的 URL 下载会通过 `AbortController` 取消；自定义异步函数本身无法被强制取消，但其过期结果不会覆盖新文件。
- 用户手动选择文件时，同样会终止当前远程下载并使旧解析结果失效。

## Props

| 属性                 | 类型                                             | 默认值         | 说明                                     |
| -------------------- | ------------------------------------------------ | -------------- | ---------------------------------------- |
| `uri`                | `OfficeViewerUri`                                | -              | 预加载文件来源                           |
| `defaultFileName`    | `string`                                         | `'未加载文件'` | 未选择文件时显示的文件名                 |
| `defaultPreviewKind` | `'pptx' \| 'xlsx' \| 'docx' \| 'doc'`            | `'pptx'`       | 无文件时使用的默认预览类型               |
| `defaultZoom`        | `number`                                         | `100`          | 默认缩放百分比，最终限制在 `25` 至 `300` |
| `className`          | `string`                                         | -              | 根容器自定义类名                         |
| `style`              | `CSSProperties`                                  | -              | 根容器自定义样式                         |
| `onFileParsed`       | `(parsed: ParsedOfficeFile, file: File) => void` | -              | 文件解析成功回调                         |
| `onError`            | `(error: Error, file?: File) => void`            | -              | 下载、解析或全屏操作失败回调             |

公开导出的类型包括：

```ts
import type {
  OfficeViewerProps,
  OfficeViewerUri,
  ParsedOfficeFile,
  PreviewKind,
} from 'office-x-viewer';
```

## 支持格式

| 格式               | 扩展名          | 支持程度     | 主要能力                                                                  |
| ------------------ | --------------- | ------------ | ------------------------------------------------------------------------- |
| PowerPoint         | `.pptx`         | 主要能力支持 | 母版与布局继承、文本、形状、图片、表格、背景、渐变、阴影、Office/WPS 图表 |
| Excel              | `.xlsx`         | 主要能力支持 | 多工作表、单元格内容与样式、合并单元格、行列尺寸、浮动图片和图表          |
| Word               | `.docx`         | 主要能力支持 | 富文本段落、表格、图片、图表、VML/WPG 形状、超链接、样式与主题色          |
| Word 97-2003 / WPS | `.doc` / `.wps` | 降级预览     | OLE/CFB、FIB、Piece Table、FKP、SPRM、PNG/JPEG 提取和纯文本回退           |

PPTX/WPS 扩展图表当前覆盖线图、柱图、饼图、环形图、面积图、散点图、气泡图、雷达图和地图等常见类型。非标准扩展或损坏数据会尽量使用文档内快照，无法降级时显示明确空状态。

## 交互说明

- 缩放范围为 `25%` 至 `300%`，工具栏快捷档位为 `50%`、`75%`、`100%`、`125%`、`150%`、`200%`。
- PPTX 支持幻灯片翻页和缩略图导航。
- XLSX 支持工作表标签切换。
- 全屏按钮依赖浏览器 Fullscreen API；不支持时按钮会禁用，按 `Esc` 退出后状态会自动同步。
- 地图图表可能需要加载外部 GeoJSON；网络失败时优先显示文档快照，否则显示加载失败状态。

## 使用边界

- 组件面向现代浏览器，依赖 `File`、`fetch`、`DOMParser`、`AbortController`、`IntersectionObserver`、`ResizeObserver` 和 Fullscreen API 等浏览器能力。
- 所有解析均在主机浏览器内完成，大文件会占用较多内存和 CPU。
- ZIP 读取仅限制并发数，不限制输入文件大小、条目数量、单个条目或累计解压大小。面向不可信文件时，建议由宿主或服务端增加业务侧校验。
- DOC/WPS 属于旧二进制格式，目前以内容可读为目标，不保证复杂分页、锚点、图文环绕和排版与 Office 完全一致。
- OOXML 文档也可能使用未覆盖的厂商扩展，复杂效果与 Microsoft Office/WPS 的原生渲染可能存在差异。

## 项目结构

```text
src/
├── index.ts
└── office-viewer/
    ├── OfficeViewer.tsx      # 对外主组件与文件加载编排
    ├── shell/                # 工具栏、预览分发和通用状态
    ├── services/             # DOC、DOCX、PPTX、XLSX 解析器
    ├── formats/              # 各文档格式的 React 渲染器
    └── shared/
        ├── ooxml/            # ZIP、XML、关系、主题、媒体和图表适配
        │   └── wpsChart.ts   # DOCX/PPTX 共用的 WPS 图表转换
        └── chart/            # ECharts 渲染和失败降级
```

数据流：

```text
文件来源
  → 格式检测
  → 对应解析器
  → TypeScript 文档模型
  → 对应 React 渲染器
  → 浏览器页面
```

## 本地开发

```bash
yarn
yarn start
```

构建组件库和文档：

```bash
yarn build
yarn docs:build
```

项目当前使用 ESLint、Stylelint、Prettier 和 TypeScript 进行静态检查。发布前建议同时在 antd v4、v5、v6 宿主项目中验证本地文件、远程 URI、全屏和各格式示例文档。
