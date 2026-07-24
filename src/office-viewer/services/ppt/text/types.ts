import type { TextParagraph, TextStyle } from '../../presentation/types';
import type { PptRecord } from '../types';

export type PptTextDefaults = {
  document?: TextStyle;
  master?: TextStyle;
  placeholder?: TextStyle;
  fonts?: Map<number, string>;
};

export type PptTextAtomGroup = {
  textType: number;
  text: string;
  contentRecord: PptRecord;
  styleRecord?: PptRecord;
};

export type PptParagraphStyleRun = {
  count: number;
  level: number;
  style: TextStyle;
};

export type PptCharacterStyleRun = {
  count: number;
  style: TextStyle;
};

export type PptTextStyleRuns = {
  paragraphs: PptParagraphStyleRun[];
  characters: PptCharacterStyleRun[];
};

export type PptParsedText = {
  textType: number;
  paragraphs: TextParagraph[];
};
