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
};
