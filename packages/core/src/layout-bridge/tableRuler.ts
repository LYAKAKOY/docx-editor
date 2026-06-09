import type { EditorState } from 'prosemirror-state';

import { getTableContext } from '../prosemirror/extensions/nodes/TableExtension';
import { MIN_CELL_WIDTH_TWIPS, type TableRulerBoundary } from '../prosemirror/tableResize';
import { pixelsToTwips } from '../utils/units';

export interface TableRulerState {
  tablePmStart: number;
  boundaries: TableRulerBoundary[];
  activeCellArea: TableRulerCellArea | null;
}

export interface TableRulerCellArea {
  leftTwips: number;
  rightTwips: number;
}

function directTableHandles(tableEl: HTMLElement, selector: string): HTMLElement[] {
  return Array.from(tableEl.querySelectorAll<HTMLElement>(selector)).filter(
    (el) => el.closest('.layout-table') === tableEl
  );
}

function elementCenterXTwips(el: HTMLElement, pageRect: DOMRect, zoom: number): number {
  const rect = el.getBoundingClientRect();
  const zoomFactor = zoom || 1;
  return Math.round(pixelsToTwips((rect.left + rect.width / 2 - pageRect.left) / zoomFactor));
}

function findDirectTableCell(
  tableEl: HTMLElement,
  rowIndex: number | undefined,
  columnIndex: number | undefined
): HTMLElement | null {
  if (rowIndex === undefined || columnIndex === undefined) return null;

  for (const rowEl of Array.from(tableEl.children)) {
    if (!(rowEl instanceof HTMLElement)) continue;
    if (!rowEl.classList.contains('layout-table-row')) continue;
    if (Number(rowEl.dataset.rowIndex) !== rowIndex) continue;

    for (const cellEl of Array.from(rowEl.children)) {
      if (!(cellEl instanceof HTMLElement)) continue;
      if (!cellEl.classList.contains('layout-table-cell')) continue;
      if (Number(cellEl.dataset.columnIndex) === columnIndex) {
        return cellEl;
      }
    }
  }

  return null;
}

function resolveActiveCellArea(
  tableEl: HTMLElement,
  pageRect: DOMRect,
  zoom: number,
  rowIndex: number | undefined,
  columnIndex: number | undefined
): TableRulerCellArea | null {
  const cellEl = findDirectTableCell(tableEl, rowIndex, columnIndex);
  if (!cellEl) return null;

  const contentEl =
    cellEl.querySelector<HTMLElement>(':scope > .layout-table-cell-content') ?? cellEl;
  const rect = contentEl.getBoundingClientRect();
  const zoomFactor = zoom || 1;
  const leftTwips = Math.round(pixelsToTwips((rect.left - pageRect.left) / zoomFactor));
  const rightTwips = Math.round(pixelsToTwips((rect.right - pageRect.left) / zoomFactor));

  if (rightTwips <= leftTwips) return null;
  return { leftTwips, rightTwips };
}

export function resolveTableRulerState(
  pagesContainer: HTMLElement | null | undefined,
  state: EditorState | null | undefined,
  options: { zoom?: number; scope?: 'body' | 'header' | 'footer' } = {}
): TableRulerState | null {
  if (!pagesContainer || !state) return null;

  const context = getTableContext(state);
  if (!context.isInTable || context.tablePos === undefined) return null;

  const scope = options.scope ?? 'body';
  const scopeClass = scope === 'body' ? 'layout-page-content' : `layout-page-${scope}`;
  const tableEl = pagesContainer.querySelector<HTMLElement>(
    `.${scopeClass} .layout-table[data-pm-start="${context.tablePos}"]`
  );
  if (!tableEl) return null;

  const pageEl = tableEl.closest<HTMLElement>('.layout-page');
  if (!pageEl) return null;

  const zoom = options.zoom || 1;
  const pageRect = pageEl.getBoundingClientRect();
  const tableRect = tableEl.getBoundingClientRect();
  const tableLeftTwips = Math.round(pixelsToTwips((tableRect.left - pageRect.left) / zoom));
  const pageWidthTwips = Math.round(pixelsToTwips(pageRect.width / zoom));

  const handles = [
    ...directTableHandles(tableEl, '.layout-table-resize-handle').map((el) => ({
      el,
      kind: 'column' as const,
    })),
    ...directTableHandles(tableEl, '.layout-table-edge-handle-right').map((el) => ({
      el,
      kind: 'rightEdge' as const,
    })),
  ]
    .map(({ el, kind }) => ({
      el,
      kind,
      columnIndex: Number(el.dataset.columnIndex),
      positionTwips: elementCenterXTwips(el, pageRect, zoom),
    }))
    .filter((handle) => Number.isFinite(handle.columnIndex))
    .sort((a, b) => a.positionTwips - b.positionTwips);

  if (handles.length === 0) return null;

  const boundaries: TableRulerBoundary[] = handles.map((handle, index) => {
    const previousTwips = index === 0 ? tableLeftTwips : handles[index - 1]!.positionTwips;
    const nextTwips = handles[index + 1]?.positionTwips;
    const minTwips = previousTwips + MIN_CELL_WIDTH_TWIPS;
    const maxTwips =
      handle.kind === 'column' && nextTwips !== undefined
        ? nextTwips - MIN_CELL_WIDTH_TWIPS
        : pageWidthTwips;

    return {
      tablePmStart: context.tablePos!,
      columnIndex: handle.columnIndex,
      kind: handle.kind,
      positionTwips: handle.positionTwips,
      previousTwips,
      nextTwips,
      minTwips,
      maxTwips: Math.max(minTwips, maxTwips),
    };
  });

  return {
    tablePmStart: context.tablePos,
    boundaries,
    activeCellArea: resolveActiveCellArea(
      tableEl,
      pageRect,
      zoom,
      context.rowIndex,
      context.columnIndex
    ),
  };
}
