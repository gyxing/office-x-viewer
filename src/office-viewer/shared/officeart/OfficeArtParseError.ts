/** 表示 OfficeArt 结构本身损坏或越过父容器边界。 */
export class OfficeArtParseError extends Error {
  readonly offset?: number;

  constructor(message: string, offset?: number) {
    super(message);
    this.name = 'OfficeArtParseError';
    this.offset = offset;
  }
}
