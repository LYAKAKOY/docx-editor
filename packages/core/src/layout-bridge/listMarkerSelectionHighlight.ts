import type { EditorState } from 'prosemirror-state';

import { resolveListMarkerSelectionGroup } from '../prosemirror/listMarkerSelection';

export function applyListMarkerSelectionHighlight(
  pagesContainer: HTMLElement,
  state: EditorState,
  options: { scope?: 'body' | 'header' | 'footer'; markerPmStart?: number | null } = {}
): void {
  const scope = options.scope ?? 'body';
  const scopeClass = scope === 'body' ? 'layout-page-content' : `layout-page-${scope}`;

  const prevSelected = pagesContainer.querySelectorAll(
    `.${scopeClass} .layout-list-marker-selected`
  );
  for (const el of Array.from(prevSelected)) {
    el.classList.remove('layout-list-marker-selected');
  }

  const selectedGroup = resolveListMarkerSelectionGroup(state.doc, options.markerPmStart);
  if (!selectedGroup) return;

  const markers = pagesContainer.querySelectorAll(`.${scopeClass} .layout-list-marker`);
  for (const marker of Array.from(markers)) {
    const htmlEl = marker as HTMLElement;
    const pmStart = htmlEl.dataset.pmStart;
    if (pmStart === undefined) continue;
    if (selectedGroup.paragraphStarts.has(Number(pmStart))) {
      htmlEl.classList.add('layout-list-marker-selected');
    }
  }
}
