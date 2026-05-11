import { attr, childByLocalName, childrenByLocalName, parseXml } from './xml';

export type OfficeTheme = {
  colorScheme: Record<string, string>;
  colorMap?: Record<string, string>;
  fontScheme?: Record<string, string>;
};

const DEFAULT_COLOR_MAP: Record<string, string> = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

export const DEFAULT_OFFICE_THEME: OfficeTheme = {
  colorMap: DEFAULT_COLOR_MAP,
  colorScheme: {
    dk1: '000000',
    lt1: 'FFFFFF',
    dk2: '44546A',
    lt2: 'E7E6E6',
    accent1: '5B9BD5',
    accent2: 'ED7D31',
    accent3: 'A5A5A5',
    accent4: 'FFC000',
    accent5: '4472C4',
    accent6: '70AD47',
    hlink: '0563C1',
    folHlink: '954F72',
  },
};

function toHexColor(value?: string) {
  if (!value) return undefined;
  if (value.startsWith('#')) return value;
  if (/^[0-9a-f]{6}$/i.test(value)) return `#${value}`;
  return undefined;
}

function readFontScheme(fontSchemeNode: Element | null | undefined) {
  const scheme: Record<string, string> = {};
  if (!fontSchemeNode) return scheme;

  ['majorFont', 'minorFont'].forEach((bucket) => {
    const node = childByLocalName(fontSchemeNode, bucket);
    if (!node) return;
    const latin = childByLocalName(node, 'latin');
    const ea = childByLocalName(node, 'ea');
    const cs = childByLocalName(node, 'cs');
    const eastAsia =
      attr(ea, 'typeface') ||
      childrenByLocalName(node, 'font').find((font) =>
        ['Hans', 'Hant', 'Jpan', 'Hang'].includes(attr(font, 'script') ?? ''),
      )?.getAttribute('typeface');
    const value = [eastAsia, attr(latin, 'typeface'), attr(cs, 'typeface')]
      .filter(Boolean)
      .join(', ');
    if (value) scheme[bucket] = value;
  });

  return scheme;
}

export function readOfficeTheme(xml?: string): OfficeTheme {
  if (!xml) return DEFAULT_OFFICE_THEME;

  const doc = parseXml(xml);
  const colorScheme: Record<string, string> = { ...DEFAULT_OFFICE_THEME.colorScheme };
  const themeElements = childByLocalName(doc.documentElement, 'themeElements');
  const clrScheme = childByLocalName(themeElements, 'clrScheme');
  const fontScheme = readFontScheme(childByLocalName(themeElements, 'fontScheme'));

  Object.keys(DEFAULT_OFFICE_THEME.colorScheme).forEach((name) => {
    const node = childByLocalName(clrScheme, name);
    const child = node?.firstElementChild;
    const value = attr(child, 'val') ?? attr(child, 'lastClr');
    if (value) colorScheme[name] = value;
  });

  return {
    colorScheme,
    colorMap: DEFAULT_COLOR_MAP,
    fontScheme,
  };
}

export function resolveOfficeThemeColor(value: string | undefined, theme: OfficeTheme = DEFAULT_OFFICE_THEME) {
  if (!value) return undefined;

  const direct = toHexColor(value);
  if (direct) return direct;

  const mapped = theme.colorMap?.[value] ?? value;
  return toHexColor(theme.colorScheme[mapped] ?? theme.colorScheme[value]);
}
