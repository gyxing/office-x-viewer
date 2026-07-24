import {
  createPptResourceId,
  registerPptResource,
  type PptParseContext,
} from '../types';

function escapeXml(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** 为无法原生渲染的嵌入对象生成稳定、可辨识的 SVG 静态预览卡片。 */
export function createPptStaticPreviewCard(
  title: string,
  detail: string,
  context: PptParseContext,
) {
  const safeTitle = escapeXml(title);
  const safeDetail = escapeXml(detail);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">
  <rect width="640" height="360" rx="24" fill="#F8FAFC"/>
  <rect x="1" y="1" width="638" height="358" rx="23" fill="none" stroke="#CBD5E1" stroke-width="2"/>
  <rect x="56" y="66" width="112" height="132" rx="14" fill="#E2E8F0"/>
  <path d="M84 98h56v68H84zM96 82h32l20 20v80H76V82z" fill="#64748B"/>
  <path d="M128 82v20h20" fill="none" stroke="#F8FAFC" stroke-width="6"/>
  <text x="200" y="128" fill="#0F172A" font-family="Arial, sans-serif" font-size="30" font-weight="700">${safeTitle}</text>
  <text x="200" y="174" fill="#475569" font-family="Arial, sans-serif" font-size="20">${safeDetail}</text>
  <rect x="200" y="212" width="310" height="12" rx="6" fill="#CBD5E1"/>
  <rect x="200" y="240" width="230" height="12" rx="6" fill="#E2E8F0"/>
</svg>`;
  return registerPptResource(context, {
    id: createPptResourceId(context, 'preview'),
    encoding: 'text',
    mimeType: 'image/svg+xml',
    text: svg,
  });
}
