import type { Node as ProseMirrorNode } from 'prosemirror-model';

type NumPrAttrs = { numId?: number | string | null; ilvl?: number | null };

export interface ListMarkerSelectionGroup {
  paragraphs: Array<{ node: ProseMirrorNode; pos: number }>;
  paragraphStarts: Set<number>;
}

function getNumPr(node: ProseMirrorNode): NumPrAttrs | null {
  const numPr = node.attrs?.numPr as NumPrAttrs | null | undefined;
  return numPr?.numId != null ? numPr : null;
}

function sameListLevel(node: ProseMirrorNode, target: NumPrAttrs): boolean {
  const numPr = getNumPr(node);
  if (!numPr) return false;
  return String(numPr.numId) === String(target.numId) && (numPr.ilvl ?? 0) === (target.ilvl ?? 0);
}

export function resolveListMarkerSelectionGroup(
  doc: ProseMirrorNode,
  markerPmStart: number | null | undefined
): ListMarkerSelectionGroup | null {
  if (markerPmStart == null || !Number.isFinite(markerPmStart)) return null;

  const paragraphs: Array<{ node: ProseMirrorNode; pos: number }> = [];
  let anchorIndex = -1;

  doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph') return;
    const index = paragraphs.length;
    paragraphs.push({ node, pos });
    if (pos === markerPmStart || (pos < markerPmStart && markerPmStart < pos + node.nodeSize)) {
      anchorIndex = index;
    }
  });

  if (anchorIndex < 0) return null;
  const anchor = paragraphs[anchorIndex];
  const target = getNumPr(anchor.node);
  if (!target) return null;

  let start = anchorIndex;
  while (start > 0 && sameListLevel(paragraphs[start - 1].node, target)) {
    start--;
  }

  let end = anchorIndex;
  while (end + 1 < paragraphs.length && sameListLevel(paragraphs[end + 1].node, target)) {
    end++;
  }

  const selected = paragraphs.slice(start, end + 1);
  return {
    paragraphs: selected,
    paragraphStarts: new Set(selected.map((p) => p.pos)),
  };
}
