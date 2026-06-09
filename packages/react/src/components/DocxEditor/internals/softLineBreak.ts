import type { EditorView } from 'prosemirror-view';

export function insertSoftLineBreak(view: EditorView): boolean {
  const hardBreakType = view.state.schema.nodes.hardBreak;
  if (!hardBreakType) return false;

  const marks = view.state.storedMarks ?? view.state.selection.$from.marks();

  try {
    const tr = view.state.tr
      .replaceSelectionWith(hardBreakType.create(null, null, marks), false)
      .ensureMarks(marks)
      .scrollIntoView();

    if (!tr.docChanged) return false;
    view.dispatch(tr);
    return true;
  } catch {
    return false;
  }
}

export function handleSoftLineBreakKey(view: EditorView, event: KeyboardEvent): boolean {
  if (event.key !== 'Enter') return false;
  if (!event.shiftKey || event.altKey || event.ctrlKey || event.metaKey) return false;
  if (event.isComposing) return false;

  return insertSoftLineBreak(view);
}
