import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Command } from 'prosemirror-state';

import type { ParagraphFormatting, TabStop } from '../../types/document';
import { resolveListMarkerSelectionGroup } from '../listMarkerSelection';

export const DEFAULT_LIST_LEFT_INDENT_TWIPS = 720;
export const DEFAULT_LIST_HANGING_TWIPS = 360;

export type RulerParagraphIndents = {
  indentLeft: number;
  indentRight: number;
  firstLineIndent: number;
  hangingIndent: boolean;
  tabStops: TabStop[] | null;
};

type ParagraphIndentAttrs = {
  numPr?: { numId?: number; ilvl?: number } | null;
  indentLeft?: number | null;
  indentRight?: number | null;
  indentFirstLine?: number | null;
  hangingIndent?: boolean | null;
  tabs?: TabStop[] | null;
};

function isListParagraphAttrs(attrs: ParagraphIndentAttrs | null | undefined): boolean {
  return !!attrs?.numPr;
}

function listLevel(attrs: ParagraphIndentAttrs): number {
  const raw = attrs.numPr?.ilvl ?? 0;
  return Number.isFinite(raw) && raw >= 0 ? raw : 0;
}

function defaultListLeftIndent(attrs: ParagraphIndentAttrs): number {
  return (listLevel(attrs) + 1) * DEFAULT_LIST_LEFT_INDENT_TWIPS;
}

function positiveOrNull(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : null;
}

function effectiveListLeftIndent(attrs: ParagraphIndentAttrs): number {
  return positiveOrNull(attrs.indentLeft) ?? defaultListLeftIndent(attrs);
}

function effectiveListHanging(attrs: ParagraphIndentAttrs): number {
  const explicit = positiveOrNull(Math.abs(attrs.indentFirstLine ?? 0));
  if (attrs.hangingIndent || explicit !== null) return explicit ?? DEFAULT_LIST_HANGING_TWIPS;
  return DEFAULT_LIST_HANGING_TWIPS;
}

export function getRulerIndentsFromParagraphFormatting(
  formatting: ParagraphFormatting | Record<string, never> | undefined
): RulerParagraphIndents {
  const pf = formatting ?? {};

  if (isListParagraphAttrs(pf)) {
    return {
      indentLeft: effectiveListLeftIndent(pf),
      indentRight: pf.indentRight ?? 0,
      firstLineIndent: effectiveListHanging(pf),
      hangingIndent: true,
      tabStops: pf.tabs ?? null,
    };
  }

  return {
    indentLeft: pf.indentLeft ?? 0,
    indentRight: pf.indentRight ?? 0,
    firstLineIndent: pf.indentFirstLine ?? 0,
    hangingIndent: pf.hangingIndent ?? false,
    tabStops: pf.tabs ?? null,
  };
}

function selectedListParagraphs(
  doc: ProseMirrorNode,
  from: number,
  to: number,
  markerPmStart?: number | null
) {
  const markerGroup = resolveListMarkerSelectionGroup(doc, markerPmStart);
  if (markerGroup) return markerGroup.paragraphs;

  const paragraphs: Array<{ node: ProseMirrorNode; pos: number }> = [];
  const seen = new Set<number>();

  doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name !== 'paragraph' || seen.has(pos) || !isListParagraphAttrs(node.attrs)) {
      return;
    }
    seen.add(pos);
    paragraphs.push({ node, pos });
  });

  return paragraphs;
}

export function setListTextIndentFromRuler(
  indentLeftTwips: number,
  markerPmStart?: number | null
): Command {
  return (state, dispatch) => {
    const { from, to } = state.selection;
    const paragraphs = selectedListParagraphs(state.doc, from, to, markerPmStart);
    if (paragraphs.length === 0) return false;
    if (!dispatch) return true;

    let tr = state.tr;
    for (const { node, pos } of paragraphs) {
      const hanging = effectiveListHanging(node.attrs);
      tr = tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        indentLeft: indentLeftTwips > 0 ? indentLeftTwips : null,
        indentFirstLine: hanging,
        hangingIndent: true,
      });
    }

    dispatch(tr.scrollIntoView());
    return true;
  };
}

export function setListMarkerIndentFromRuler(
  firstLineTwips: number,
  markerPmStart?: number | null
): Command {
  return (state, dispatch) => {
    const { from, to } = state.selection;
    const paragraphs = selectedListParagraphs(state.doc, from, to, markerPmStart);
    if (paragraphs.length === 0) return false;
    if (!dispatch) return true;

    let tr = state.tr;
    for (const { node, pos } of paragraphs) {
      const indentLeft = effectiveListLeftIndent(node.attrs);
      const desiredMarkerPos = indentLeft + firstLineTwips;
      const desiredGap = indentLeft - desiredMarkerPos;
      const nextGap = Math.max(DEFAULT_LIST_HANGING_TWIPS, desiredGap);
      const nextIndentLeft =
        desiredGap < DEFAULT_LIST_HANGING_TWIPS
          ? Math.max(0, desiredMarkerPos + DEFAULT_LIST_HANGING_TWIPS)
          : indentLeft;

      tr = tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        indentLeft: nextIndentLeft > 0 ? nextIndentLeft : null,
        indentFirstLine: nextGap,
        hangingIndent: true,
      });
    }

    dispatch(tr.scrollIntoView());
    return true;
  };
}
