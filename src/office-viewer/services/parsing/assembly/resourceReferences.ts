const RESOURCE_REFERENCE_PREFIX = 'office-resource:';

/** 创建只在解析传输模型内部使用的资源引用。 */
export function createResourceReference(id: string) {
  return `${RESOURCE_REFERENCE_PREFIX}${id}`;
}

/** 从解析传输模型的资源引用中读取稳定资源 ID。 */
export function readResourceReference(value: string) {
  return value.startsWith(RESOURCE_REFERENCE_PREFIX)
    ? value.slice(RESOURCE_REFERENCE_PREFIX.length)
    : undefined;
}
