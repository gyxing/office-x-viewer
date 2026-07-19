import type { ReactNode, SVGProps } from 'react';
import React from 'react';

type OfficeIconProps = Omit<SVGProps<SVGSVGElement>, 'children'>;

type OfficeIconBaseProps = OfficeIconProps & {
  children: ReactNode;
};

const FILE_FRAME = <path d="M6 2.75h7l5 5v13.5H6zM13 2.75v5h5" />;

// OfficeIconBase 统一工具栏图标的尺寸、描边和无障碍属性。
function OfficeIconBase({ children, ...props }: OfficeIconBaseProps) {
  return (
    <svg
      {...props}
      aria-hidden="true"
      focusable="false"
      width="1em"
      height="1em"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      {children}
    </svg>
  );
}

// FileExcelIcon 表示电子表格文件。
export function FileExcelIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      {FILE_FRAME}
      <path d="M8.5 11h7v6h-7zM12 11v6M8.5 14h7" />
    </OfficeIconBase>
  );
}

// FilePptIcon 表示演示文稿文件。
export function FilePptIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      {FILE_FRAME}
      <path d="M9 16.5v-5h2.5a1.75 1.75 0 0 1 0 3.5H9" />
    </OfficeIconBase>
  );
}

// FileWordIcon 表示文字文档文件。
export function FileWordIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      {FILE_FRAME}
      <path d="m8.5 11 1.25 6 2.25-4 2.25 4 1.25-6" />
    </OfficeIconBase>
  );
}

// ChevronLeftIcon 表示上一页操作。
export function ChevronLeftIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="m15 18-6-6 6-6" />
    </OfficeIconBase>
  );
}

// ChevronRightIcon 表示下一页操作。
export function ChevronRightIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="m9 6 6 6-6 6" />
    </OfficeIconBase>
  );
}

// ZoomOutIcon 表示缩小操作。
export function ZoomOutIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <circle cx="10.5" cy="10.5" r="6.5" />
      <path d="M7.5 10.5h6M15.5 15.5 21 21" />
    </OfficeIconBase>
  );
}

// ZoomInIcon 表示放大操作。
export function ZoomInIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <circle cx="10.5" cy="10.5" r="6.5" />
      <path d="M7.5 10.5h6M10.5 7.5v6M15.5 15.5 21 21" />
    </OfficeIconBase>
  );
}

// FullscreenIcon 表示进入全屏操作。
export function FullscreenIcon(props: OfficeIconProps) {
  return (
    <OfficeIconBase {...props}>
      <path d="M8 3H3v5M16 3h5v5M21 16v5h-5M8 21H3v-5" />
    </OfficeIconBase>
  );
}
