export function buildRendererId(renderKey: string, elementId: string, suffix: string) {
  return `${renderKey}-${elementId}-${suffix}`.replace(/[^a-zA-Z0-9_-]/g, '-');
}

