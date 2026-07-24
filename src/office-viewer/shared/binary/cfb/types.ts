export type CfbObjectType = 'storage' | 'stream' | 'root';

export type CfbDirectoryEntry = {
  id: number;
  name: string;
  path: string;
  objectType: CfbObjectType;
  startSector: number;
  streamSize: number;
  leftSiblingId: number;
  rightSiblingId: number;
  childId: number;
};

export type CfbFile = {
  entries: CfbDirectoryEntry[];
  streams: Map<string, Uint8Array>;
  getStream: (...names: string[]) => Uint8Array | undefined;
  hasEntry: (name: string) => boolean;
};

export type CfbReadOptions = {
  yieldIfNeeded?: () => Promise<void>;
  /** 兼容省略最后扇区零填充的 CFB 生成器；默认保持严格校验。 */
  allowPartialFinalSector?: boolean;
};
