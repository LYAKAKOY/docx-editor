import { test, expect, type Locator, type Page } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';

async function dragMarker(page: Page, marker: Locator, deltaX: number): Promise<void> {
  const box = await marker.boundingBox();
  expect(box).not.toBeNull();
  if (!box) return;

  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + deltaX, y, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(150);
}

async function paragraphAttrs(page: Page, index: number): Promise<Record<string, unknown>> {
  const attrs = await page.evaluate((i) => window.__DOCX_EDITOR_E2E__?.getParagraphAttrs(i), index);
  expect(attrs).not.toBeNull();
  return attrs ?? {};
}

async function visualTableCellWidths(page: Page): Promise<number[]> {
  return page.evaluate(() => {
    const row = document.querySelector(
      '.paged-editor__pages .layout-page-content .layout-table .layout-table-row'
    );
    if (!row) return [];
    return Array.from(row.children)
      .filter(
        (el): el is HTMLElement =>
          el instanceof HTMLElement && el.classList.contains('layout-table-cell')
      )
      .map((cell) => cell.getBoundingClientRect().width);
  });
}

async function markerCenterX(marker: Locator): Promise<number> {
  const box = await marker.boundingBox();
  expect(box).not.toBeNull();
  return box ? box.x + box.width / 2 : 0;
}

function expectClosePx(actual: number, expected: number, tolerance = 8): void {
  expect(Math.abs(actual - expected)).toBeLessThanOrEqual(tolerance);
}

async function visualTableCellContentRect(
  page: Page,
  columnIndex: number
): Promise<{ left: number; right: number; width: number }> {
  const rect = await page.evaluate((col) => {
    const row = document.querySelector(
      '.paged-editor__pages .layout-page-content .layout-table .layout-table-row'
    );
    const cell = row?.querySelector<HTMLElement>(`.layout-table-cell[data-column-index="${col}"]`);
    const content = cell?.querySelector<HTMLElement>('.layout-table-cell-content') ?? cell;
    if (!content) return null;
    const box = content.getBoundingClientRect();
    return { left: box.left, right: box.right, width: box.width };
  }, columnIndex);
  expect(rect).not.toBeNull();
  return rect ?? { left: 0, right: 0, width: 0 };
}

test.describe('horizontal ruler paragraph indents', () => {
  test('shows separate first-line and left-indent handles', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    const firstLineMarker = page.getByRole('slider', { name: 'First line indent' });
    const leftIndentMarker = page.getByRole('slider', { name: 'Left indent' });
    const rightIndentMarker = page.getByRole('slider', { name: 'Right indent' });

    await expect(firstLineMarker).toBeVisible();
    await expect(leftIndentMarker).toBeVisible();
    await expect(rightIndentMarker).toBeVisible();
    await expect(page.locator('.docx-ruler-indent')).toHaveCount(3);

    const firstLineBox = await firstLineMarker.boundingBox();
    const leftIndentBox = await leftIndentMarker.boundingBox();
    const rightIndentBox = await rightIndentMarker.boundingBox();
    expect(firstLineBox).not.toBeNull();
    expect(leftIndentBox).not.toBeNull();
    expect(rightIndentBox).not.toBeNull();
    if (!firstLineBox || !leftIndentBox || !rightIndentBox) return;

    expect(leftIndentBox.y).toBeGreaterThan(firstLineBox.y);
    expect(rightIndentBox.y).toBeGreaterThan(firstLineBox.y);
  });

  test('first-line handle indents only first lines across selected paragraphs', async ({
    page,
  }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.focusParagraph(0);
    await editor.typeText('First selected paragraph wraps with enough words for ruler testing.');
    await editor.pressEnter();
    await editor.typeText('Second selected paragraph should receive the same first-line indent.');
    await editor.selectAll();

    const firstLineMarker = page.getByRole('slider', { name: 'First line indent' });
    await dragMarker(page, firstLineMarker, 48);

    for (const index of [0, 1]) {
      const attrs = await paragraphAttrs(page, index);
      expect(attrs.indentFirstLine).toBeGreaterThan(0);
      expect(attrs.hangingIndent).toBe(false);
      expect(attrs.indentLeft ?? null).toBeNull();
    }
  });

  test('left-indent handle indents the whole selected paragraph', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.focusParagraph(0);
    await editor.typeText('Whole paragraph indent from the lower ruler marker.');
    await editor.selectAll();

    const leftIndentMarker = page.getByRole('slider', { name: 'Left indent' });
    await dragMarker(page, leftIndentMarker, 48);

    const attrs = await paragraphAttrs(page, 0);
    expect(attrs.indentLeft).toBeGreaterThan(0);
    expect(attrs.indentFirstLine ?? null).toBeNull();
  });

  test('right-indent handle indents the selected paragraph from the right', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.focusParagraph(0);
    await editor.typeText('Whole paragraph right indent from the lower right ruler marker.');
    await editor.selectAll();

    const rightIndentMarker = page.getByRole('slider', { name: 'Right indent' });
    await dragMarker(page, rightIndentMarker, -48);

    const attrs = await paragraphAttrs(page, 0);
    expect(attrs.indentRight).toBeGreaterThan(0);
    expect(attrs.indentLeft ?? null).toBeNull();
    expect(attrs.indentFirstLine ?? null).toBeNull();
  });

  test('clicking a list marker highlights the marker group only after marker click', async ({
    page,
  }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.focusParagraph(0);
    await editor.typeText('привет');
    await editor.toggleNumberedList();
    await editor.pressEnter();
    await editor.typeText('пока');
    await editor.pressEnter();
    await editor.typeText('привет');
    await editor.pressEnter();
    await editor.typeText('пока');

    const markers = page.locator('.layout-page-content .layout-list-marker');
    const selectedMarkers = page.locator('.layout-page-content .layout-list-marker-selected');
    await expect(markers).toHaveCount(4);
    await expect(selectedMarkers).toHaveCount(0);

    await markers.first().click();
    await expect(selectedMarkers).toHaveCount(4);

    const firstLineMarker = page.getByRole('slider', { name: 'First line indent' });
    const leftIndentMarker = page.getByRole('slider', { name: 'Left indent' });
    const firstLineBox = await firstLineMarker.boundingBox();
    const leftIndentBox = await leftIndentMarker.boundingBox();
    expect(firstLineBox).not.toBeNull();
    expect(leftIndentBox).not.toBeNull();
    if (!firstLineBox || !leftIndentBox) return;

    expect(leftIndentBox.x - firstLineBox.x).toBeGreaterThan(10);

    await dragMarker(page, leftIndentMarker, 36);
    for (const index of [0, 1, 2, 3]) {
      const attrs = await paragraphAttrs(page, index);
      expect(attrs.indentLeft).toBeGreaterThan(720);
    }
    await expect(selectedMarkers).toHaveCount(4);

    const textRun = page
      .locator('.layout-page-content span[data-pm-start][data-pm-end]:not(.layout-list-marker)')
      .filter({ hasText: 'пока' })
      .first();
    await textRun.click();
    await expect(selectedMarkers).toHaveCount(0);
  });

  test('list first-line marker moves numbers first and then preserves the marker-text gap', async ({
    page,
  }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.focusParagraph(0);
    await editor.typeText('First list item.');
    await editor.toggleNumberedList();
    await page.locator('.layout-page-content .layout-list-marker').first().click();

    const firstLineMarker = page.getByRole('slider', { name: 'First line indent' });
    await dragMarker(page, firstLineMarker, -48);

    let attrs = await paragraphAttrs(page, 0);
    expect(attrs.indentLeft).toBeGreaterThanOrEqual(700);
    expect(attrs.indentLeft).toBeLessThanOrEqual(740);
    expect(attrs.indentFirstLine).toBeGreaterThanOrEqual(700);
    expect(attrs.indentFirstLine).toBeLessThanOrEqual(740);
    expect(attrs.hangingIndent).toBe(true);

    await dragMarker(page, firstLineMarker, 48);

    attrs = await paragraphAttrs(page, 0);
    expect(attrs.indentLeft).toBeGreaterThan(720);
    expect(attrs.indentFirstLine).toBeGreaterThanOrEqual(340);
    expect(attrs.indentFirstLine).toBeLessThanOrEqual(380);
    expect(attrs.hangingIndent).toBe(true);
  });

  test('list left-indent marker moves marker and text together', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.focusParagraph(0);
    await editor.typeText('Whole list item follows the lower marker.');
    await editor.toggleNumberedList();
    await page.locator('.layout-page-content .layout-list-marker').first().click();

    const leftIndentMarker = page.getByRole('slider', { name: 'Left indent' });
    await dragMarker(page, leftIndentMarker, 48);

    const attrs = await paragraphAttrs(page, 0);
    expect(attrs.indentLeft).toBeGreaterThan(720);
    expect(attrs.indentFirstLine).toBeGreaterThanOrEqual(340);
    expect(attrs.indentFirstLine).toBeLessThanOrEqual(380);
    expect(attrs.hangingIndent).toBe(true);
  });

  test('typing in a new list item keeps the caret on the text, not on the number', async ({
    page,
  }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.focusParagraph(0);
    await editor.typeText('привет');
    await editor.toggleNumberedList();
    await editor.pressEnter();
    await editor.typeText('юпривет');

    const secondMarker = page.locator('.layout-page-content .layout-list-marker').nth(1);
    const typedRun = page
      .locator('.layout-page-content span[data-pm-start][data-pm-end]:not(.layout-list-marker)')
      .filter({ hasText: 'юпривет' })
      .last();
    const caret = page.getByTestId('caret');

    await expect(secondMarker).toBeVisible();
    await expect(typedRun).toBeVisible();
    await expect(caret).toBeVisible();

    const markerBox = await secondMarker.boundingBox();
    const runBox = await typedRun.boundingBox();
    const caretBox = await caret.boundingBox();
    expect(markerBox).not.toBeNull();
    expect(runBox).not.toBeNull();
    expect(caretBox).not.toBeNull();
    if (!markerBox || !runBox || !caretBox) return;

    expect(caretBox.x).toBeGreaterThan(markerBox.x + markerBox.width + 4);
    expect(caretBox.x).toBeGreaterThan(runBox.x);
    expect(caretBox.x).toBeLessThan(runBox.x + runBox.width + 16);
  });
});

test.describe('horizontal ruler table handles', () => {
  test('aligns paragraph indent handles to the active table cell', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.insertTable(2, 3);
    await editor.clickTableCell(0, 0, 0);

    const firstLineMarker = page.getByRole('slider', { name: 'First line indent' });
    const leftIndentMarker = page.getByRole('slider', { name: 'Left indent' });
    const rightIndentMarker = page.getByRole('slider', { name: 'Right indent' });
    const firstCell = await visualTableCellContentRect(page, 0);

    await expect(firstLineMarker).toBeVisible();
    await expect(leftIndentMarker).toBeVisible();
    await expect(rightIndentMarker).toBeVisible();

    expectClosePx(await markerCenterX(firstLineMarker), firstCell.left);
    expectClosePx(await markerCenterX(leftIndentMarker), firstCell.left);
    expectClosePx(await markerCenterX(rightIndentMarker), firstCell.right);

    await editor.clickTableCell(0, 0, 1);
    const secondCell = await visualTableCellContentRect(page, 1);

    const firstLineX = await markerCenterX(firstLineMarker);
    const leftIndentX = await markerCenterX(leftIndentMarker);
    const rightIndentX = await markerCenterX(rightIndentMarker);

    expect(firstLineX).toBeGreaterThan(firstCell.right - 8);
    expectClosePx(firstLineX, secondCell.left);
    expectClosePx(leftIndentX, secondCell.left);
    expectClosePx(rightIndentX, secondCell.right);
  });

  test('shows table boundary markers and resizes adjacent columns', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.insertTable(2, 3);
    await editor.clickTableCell(0, 0, 0);

    const markers = page.locator('.docx-ruler-table-boundary');
    await expect(markers).toHaveCount(3);
    await expect(page.getByRole('slider', { name: 'Table column boundary' }).first()).toBeVisible();

    const before = await visualTableCellWidths(page);
    expect(before).toHaveLength(3);

    await dragMarker(page, markers.first(), 45);

    const after = await visualTableCellWidths(page);
    expect(after).toHaveLength(3);
    expect(after[0]).toBeGreaterThan(before[0] + 10);
    expect(after[1]).toBeLessThan(before[1] - 10);
  });

  test('right-edge table marker resizes the last column', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();

    await editor.insertTable(2, 2);
    await editor.clickTableCell(0, 0, 0);

    const markers = page.locator('.docx-ruler-table-boundary');
    await expect(markers).toHaveCount(2);

    const beforeCells = await visualTableCellWidths(page);
    expect(beforeCells).toHaveLength(2);

    await dragMarker(page, markers.last(), 45);

    const afterCells = await visualTableCellWidths(page);
    expect(afterCells).toHaveLength(2);
    expect(afterCells[1]).toBeGreaterThan(beforeCells[1] + 10);
  });
});
