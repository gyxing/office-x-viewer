# Office Viewer

纯前端、浏览器端的 Microsoft Office 文档预览组件。无需服务端参与，所有文件解析和渲染均在客户端完成。基于 React + Ant Design 构建。

## 支持格式

| 格式 | 扩展名 | 支持程度 | 说明 |
|------|--------|---------|------|
| PowerPoint | `.pptx` | 完整支持 | 幻灯片母版/布局占位符继承、文本、形状（预设/自定义几何）、图片（裁剪/旋转/翻转）、表格、图表（含 WPS 扩展图表：线/柱/饼/环/面积/散点/气泡/雷达/地图）、渐变填充、阴影、背景图/填充 |
| Excel | `.xlsx` | 完整支持 | 多工作表切换、单元格值（共享字符串/内联字符串/布尔）、单元格样式（字体/填充/边框/对齐）、合并单元格、列宽行高、浮动图片、浮动图表 |
| Word | `.docx` | 完整支持 | 段落（富文本）、表格、内联图片（DrawingML）、内联图表、VML 形状（矩形/椭圆/自定义路径）、WPG 组合形状、超链接、段落边框/底纹、样式继承、主题色解析 |
| Word 97-2003 | `.doc` | 降级预览 | OLE/CFB 二进制格式解析（FIB/Piece Table/FKP/SPRM）、嵌入图片提取（PNG/JPEG）、纯文本回退。复杂排版和图文定位为近似效果 |

## 基本用法

```tsx
import { OfficeViewer } from './office-viewer';

// 基础用法 - 用户手动选择本地文件
<OfficeViewer />

// 预设文件或远程地址
<OfficeViewer uri={file} />
<OfficeViewer uri="https://example.com/demo.pptx" />

// 懒加载文件来源
<OfficeViewer uri={async () => fetch('/demo.xlsx')} />

// 带回调
<OfficeViewer
  onFileParsed={(parsed, file) => console.log('解析完成', parsed.kind)}
  onError={(err) => console.error(err)}
/>
```

## Props

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `uri` | `File \| string \| (() => Promise<File \| Blob \| string \| Response>)` | - | 预加载文件来源。`File` 直接解析，`string` 视为 URL 下载后解析，函数用于懒加载文件来源 |
| `defaultFileName` | `string` | `'未加载文件'` | 未选择文件时显示的文件名 |
| `defaultPreviewKind` | `PreviewKind` | `'pptx'` | 默认预览格式（无文件时） |
| `defaultZoom` | `number` | `100` | 默认缩放比例 |
| `className` | `string` | - | 自定义类名 |
| `style` | `CSSProperties` | - | 自定义样式 |
| `onFileParsed` | `(parsed, file) => void` | - | 文件解析成功回调 |
| `onError` | `(error, file?) => void` | - | 解析失败回调 |

## 目录结构

```
src/office-viewer/
├── index.ts                    # 模块入口，导出 OfficeViewer 组件和类型
├── index.less                  # 全局布局样式
├── OfficeViewer.tsx            # 主编排组件，管理文件状态/解析/缩放/翻页
│
├── shell/                      # 应用外壳（所有格式共用的 UI 框架）
│   ├── PreviewStage.tsx        #   路由分发，懒加载对应格式的 Viewer
│   ├── Toolbar.tsx             #   顶部工具栏：文件上传/翻页/缩放/全屏
│   ├── Loading.tsx             #   加载中状态（居中 Spinner）
│   ├── Error.tsx               #   错误提示
│   ├── Empty.tsx               #   空状态占位（按格式显示不同提示语）
│   └── constants.ts            #   共享常量：工具栏高度(56px)、缩放级别(50-300%)
│
├── services/                   # 解析层 - 将二进制文件解析为 TypeScript 模型
│   ├── preview.ts              #   格式检测（按扩展名）& 动态分发到对应解析器
│   ├── doc/                    #   .doc 解析器
│   │   ├── parseDoc.ts         #     OLE/CFB 二进制解析：FAT/目录/FIB/Piece Table/FKP/SPRM/字体表/嵌入图片
│   │   └── types.ts            #     DOC 模型类型定义
│   ├── docx/                   #   .docx 解析器
│   │   ├── parseDocx.ts        #     OOXML ZIP 解析：样式/主题/段落/表格/Drawing/VML/WPS图表
│   │   ├── types.ts            #     DOCX 模型类型定义
│   │   └── archive.ts          #     共享 ZIP 加载器的重导出
│   ├── pptx/                   #   .pptx 解析器
│   │   ├── parsePptx.ts        #     OOXML ZIP 解析：母版/布局/占位符继承/视觉元素/图表/表格/WPS扩展
│   │   ├── types.ts            #     PPTX 模型类型定义
│   │   ├── colors.ts           #     OOXML 颜色工具：hex标准化/主题色解析/tint/shade/lumMod/lumOff
│   │   └── mediaBase64/        #     预编码媒体资源（保留目录）
│   └── xlsx/                   #   .xlsx 解析器
│       ├── parseXlsx.ts        #     OOXML ZIP 解析：共享字符串/样式表/单元格/合并区域/浮动图片/图表
│       ├── types.ts            #     XLSX 模型类型定义
│       └── archive.ts          #     共享 ZIP 加载器的重导出
│
├── formats/                    # 渲染层 - 将解析模型渲染为 React 组件
│   ├── doc/                    #   DOC 渲染
│   │   ├── DocViewer.tsx       #     顶层组件：标题/统计/警告/滚动页面/图片画廊
│   │   ├── DocContentRenderer.tsx  #  内容块分发：段落/表格/列表
│   │   ├── DocParagraphBlock.tsx   #  段落渲染（内联文本和图片）
│   │   ├── DocTableBlock.tsx   #     表格渲染（带样式的单元格）
│   │   ├── DocListBlock.tsx    #     列表渲染（有序/无序）
│   │   ├── DocInlineContent.tsx    #  内联元素渲染（文本样式/图片）
│   │   ├── DocImageGallery.tsx #     未锚定图片的画廊展示
│   │   ├── DocPageFrame.tsx    #     页面框架（尺寸/边距/缩放）
│   │   └── ...
│   ├── docx/                   #   DOCX 渲染
│   │   ├── DocxViewer.tsx      #     顶层组件：标题/统计/滚动页面
│   │   ├── DocxParagraph.tsx   #     段落渲染（富文本内联内容）
│   │   ├── DocxTableBlock.tsx  #     表格渲染（单元格样式/嵌套块内容）
│   │   ├── DocxImage.tsx       #     内联图片渲染
│   │   ├── DocxChartBlock.tsx  #     内联图表渲染（通过 OfficeChartView）
│   │   ├── DocxShape.tsx       #     形状渲染（VML/DrawingML，含文本框/填充/描边）
│   │   ├── DocxPageFrame.tsx   #     页面框架（尺寸/边距/边框/缩放）
│   │   └── ...
│   ├── pptx/                   #   PPTX 渲染
│   │   ├── PptxViewer.tsx      #     顶层组件：双面板布局（缩略图侧栏 + 幻灯片视口）
│   │   ├── PptxSlideViewport.tsx   # 幻灯片视口（滚动/缩放）
│   │   ├── PptxSlide.tsx       #     单张幻灯片渲染（背景 + 元素分发）
│   │   ├── PptxThumbnailPane.tsx   # 缩略图侧栏（可滚动列表 + 页码）
│   │   ├── PptxThumbnail.tsx   #     单个缩略图（复用 PptxSlide 缩小渲染）
│   │   └── renderers/          #     专用元素渲染器
│   │       ├── TextRenderer.tsx    # 文本框：形状填充/渐变/段落样式/项目符号
│   │       ├── ShapeRenderer.tsx   # 预设/自定义几何形状
│   │       ├── ImageRenderer.tsx   # 图片：裁剪/旋转/翻转变换
│   │       ├── TableRenderer.tsx   # 表格网格：单元格填充/边框/文本
│   │       ├── UnsupportedRenderer.tsx  # 不支持元素的占位显示
│   │       ├── paint.ts        #     颜色/渐变到 CSS/SVG 的转换工具
│   │       └── renderIds.ts    #     稳定 SVG ID 生成（避免缩略图/视口冲突）
│   └── xlsx/                   #   XLSX 渲染
│       ├── XlsxViewer.tsx      #     顶层组件：工作表标签 + 活动工作表网格
│       ├── XlsxSheetGrid.tsx   #     可滚动画布（表格 + 浮动图片 + 浮动图表）
│       ├── XlsxSheetTable.tsx  #     单元格网格渲染（表头/合并单元格/单元格样式）
│       ├── XlsxSheetTabs.tsx   #     底部工作表切换标签栏
│       ├── XlsxFloatingImages.tsx  # 锚定到单元格位置的浮动图片
│       ├── XlsxFloatingCharts.tsx  # 锚定到单元格位置的浮动图表
│       └── sheetRenderUtils.ts #     工作表画布尺寸计算工具
│
└── shared/                     # 共享工具
    ├── ooxml/                  #   OOXML 底层解析工具（所有格式共用）
    │   ├── archive.ts          #     JSZip 解包：返回 Map<string, string | Uint8Array>，提供 readXml/readBinary
    │   ├── xml.ts              #     DOMParser XML 工具：命名空间感知的元素遍历/属性读取/文本提取
    │   ├── units.ts            #     单位换算：EMU -> px（1英寸 = 914400 EMU = 96px）
    │   ├── relationships.ts    #     .rels 关系文件解析（链接 OOXML 各部分）
    │   ├── media.ts            #     媒体文件收集：从 ZIP 提取图片，转为 base64 data URL
    │   ├── theme.ts            #     主题 XML 解析：颜色方案/字体方案/主题色引用解析
    │   └── charts.ts           #     图表 XML 解析：OfficeChartModel 中间表示 -> ECharts 配置
    │                           #     支持：线/柱(垂直/水平)/饼/环/面积/散点/气泡/雷达/复合饼图/地图
    └── chart/                  #   图表渲染组件
        ├── OfficeChartView.tsx #     ECharts 图表组件：懒加载/IntersectionObserver 延迟渲染/地图 GeoJSON/ResizeObserver
        └── index.less          #     图表容器样式
```

## 整体架构

采用三层分离设计，数据流为：

```
文件上传 → 格式检测(services/preview.ts) → 对应解析器(services/*) → TypeScript 模型 → 对应渲染器(formats/*) → 页面展示
```

- **Shell 层** (`shell/`)：通用外壳 UI，负责工具栏交互、路由分发、加载/错误/空状态展示
- **Services 层** (`services/`)：各格式解析器，将二进制文件解析为结构化 TypeScript 模型
- **Formats 层** (`formats/`)：各格式渲染器，将模型渲染为 React 组件
- **Shared 层** (`shared/`)：OOXML 底层工具和图表渲染，被 Services 和 Formats 共用
