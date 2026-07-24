/** CFB 文件签名。 */
export const CFB_SIGNATURE = [
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
] as const;

/** CFB 扇区链中的空闲扇区标记。 */
export const FREE_SECTOR = 0xffffffff;
/** CFB 扇区链结束标记。 */
export const END_OF_CHAIN = 0xfffffffe;
/** FAT 扇区标记。 */
export const FAT_SECTOR = 0xfffffffd;
/** DIFAT 扇区标记。 */
export const DIFAT_SECTOR = 0xfffffffc;
/** 目录树中的空节点标记。 */
export const NO_STREAM = 0xffffffff;
/** 小流与常规流的固定分界。 */
export const MINI_STREAM_CUTOFF_SIZE = 4096;

export const CFB_HEADER_SIZE = 512;
export const CFB_DIRECTORY_ENTRY_SIZE = 128;
export const CFB_MINI_SECTOR_SIZE = 64;
