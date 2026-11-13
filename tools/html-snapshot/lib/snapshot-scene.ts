// The snapshot/importer contract distinguishes authored boxes from generated paint primitives.
// Keep browser helpers as named functions: their source is embedded in the browser payload.

export type SnapshotNodeRole = 'element' | 'fragment' | 'decoration' | 'text' | 'image' |
  'vector' | 'geometry' | 'clip' | 'coordinate';

export interface SnapshotSource {
  id: number;
  parentId: number | null;
  layoutParentId: number | null;
  absoluteContainingBlockId: number | null;
  fixedContainingBlockId: number | null;
  tag: string;
  elementId: string;
  pseudo: string | null;
  border: number[];
  padding: number[];
  clipIds: number[];
  stackingContext: boolean;
  geometry: { left: number; top: number; width: number; height: number } | null;
}

export interface SnapshotScene {
  version: 1;
  sources: SnapshotSource[];
}

function snapshotInsets(style: CSSStyleDeclaration, prefix: string, suffix = ''): number[] {
  return ['top', 'right', 'bottom', 'left'].map(side =>
    parseFloat(style.getPropertyValue(`${prefix}-${side}${suffix}`)) || 0);
}

function snapshotNodeRole(element: Element): SnapshotNodeRole {
  if (element.localName === 'span') return 'text';
  if (element.localName === 'img') return 'image';
  if (element.localName === 'svg') return 'vector';
  return 'geometry';
}

function markSnapshotNode(element: Element, role: SnapshotNodeRole, sourceId: number | null,
                          parentId: number | null = null): void {
  element.setAttribute('data-pagx-node-role', role);
  if (sourceId !== null) element.setAttribute('data-pagx-node-source', String(sourceId));
  if (parentId !== null) element.setAttribute('data-pagx-node-parent', String(parentId));
}

// Generated text/images belong to their source element. Their DOM tags describe the browser
// preview, not independent editable PAGX layers. SVG descendants retain SVG's own paint model.
function markSnapshotPrimitives(element: Element, sourceId: number | null): void {
  for (const child of Array.from(element.children)) {
    if (child.hasAttribute('data-pagx-paint-source')) continue;
    markSnapshotNode(child, snapshotNodeRole(child), sourceId);
    if (child.localName !== 'svg') markSnapshotPrimitives(child, sourceId);
  }
}

const SNAPSHOT_SCENE_HELPERS_SRC = [
  snapshotInsets, snapshotNodeRole, markSnapshotNode, markSnapshotPrimitives,
].map(fn => fn.toString()).join('\n\n');

export { SNAPSHOT_SCENE_HELPERS_SRC, snapshotInsets, snapshotNodeRole,
  markSnapshotNode, markSnapshotPrimitives };
