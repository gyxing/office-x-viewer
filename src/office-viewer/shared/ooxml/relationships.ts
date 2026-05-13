import { attr, descendantsByLocalName, parseXml } from './xml';
import { normalizeRelationshipTarget, type OfficeRelationship } from './media';

export function readRelationships(xml: string, relsPath: string) {
  const doc = parseXml(xml);
  const relationships: Record<string, OfficeRelationship> = {};
  descendantsByLocalName(doc.documentElement, 'Relationship').forEach((node) => {
    const id = attr(node, 'Id');
    const target = attr(node, 'Target');
    if (!id || !target) return;
    relationships[id] = {
      id,
      target: normalizeRelationshipTarget(relsPath, target),
      type: attr(node, 'Type'),
    };
  });
  return relationships;
}
