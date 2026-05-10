import { attr, childByLocalName, parseXml } from './xml';

export type OfficeTheme = {
  colorScheme: Record<string, string>;
  colorMap?: Record<string, string>;
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

export function readOfficeTheme(xml?: string): OfficeTheme {
  if (!xml) return DEFAULT_OFFICE_THEME;

  const doc = parseXml(xml);
  const colorScheme: Record<string, string> = { ...DEFAULT_OFFICE_THEME.colorScheme };
  const clrScheme = childByLocalName(childByLocalName(doc.documentElement, 'themeElements'), 'clrScheme');

  Object.keys(DEFAULT_OFFICE_THEME.colorScheme).forEach((name) => {
    const node = childByLocalName(clrScheme, name);
    const child = node?.firstElementChild;
    const value = attr(child, 'val') ?? attr(child, 'lastClr');
    if (value) colorScheme[name] = value;
  });

  return {
    colorScheme,
    colorMap: DEFAULT_COLOR_MAP,
  };
}

export function resolveOfficeThemeColor(value: string | undefined, theme: OfficeTheme = DEFAULT_OFFICE_THEME) {
  if (!value) return undefined;

  const direct = toHexColor(value);
  if (direct) return direct;

  const mapped = theme.colorMap?.[value] ?? value;
  return toHexColor(theme.colorScheme[mapped] ?? theme.colorScheme[value]);
}
