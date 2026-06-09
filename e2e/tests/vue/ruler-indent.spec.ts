import { test, expect, type Locator, type Page } from '@playwright/test';

const VUE_EMPTY_URL = 'http://localhost:5174/?e2e=1&empty=1';

async function gotoVueEmpty(page: Page): Promise<void> {
  await page.goto(VUE_EMPTY_URL);
  await page.locator('.docx-editor-vue').waitFor();
  await page.locator('.paged-editor__pages').waitFor();
  await page.waitForSelector('[data-page-number]');
}

async function markerCenterX(marker: Locator): Promise<number> {
  const box = await marker.boundingBox();
  expect(box).not.toBeNull();
  return box ? box.x + box.width / 2 : 0;
}

function expectClosePx(actual: number, expected: number, tolerance = 8): void {
  expect(Math.abs(actual - expected)).toBeLessThanOrEqual(tolerance);
}

async function insertVueTable(page: Page, rows: number, cols: number): Promise<void> {
  await page.locator('.layout-page-content').click();
  await page.getByRole('button', { name: /^Insert$/ }).click();
  await page.getByRole('button', { name: /^Table$/ }).hover();

  const grid = page.getByRole('grid', { name: 'Table size selector' });
  await grid.waitFor({ state: 'visible', timeout: 5000 });

  const gridColumns = await grid.evaluate((el) => {
    const style = window.getComputedStyle(el);
    const columns = style.gridTemplateColumns.split(/\s+/).filter(Boolean);
    return columns.length || 6;
  });
  const target = grid.getByRole('gridcell').nth((rows - 1) * gridColumns + (cols - 1));
  await target.hover();
  await page.waitForTimeout(100);
  await target.click();

  await page.waitForSelector('.layout-page-content .layout-table', {
    state: 'visible',
    timeout: 5000,
  });
}

async function clickVueTableCell(page: Page, row: number, col: number): Promise<void> {
  const cell = page
    .locator('.layout-page-content .layout-table-row')
    .nth(row)
    .locator(`.layout-table-cell[data-column-index="${col}"]`);
  await cell.click();
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

test.describe('Vue horizontal ruler indents', () => {
  test('aligns paragraph indent handles to the active table cell', async ({ page }) => {
    await gotoVueEmpty(page);

    await insertVueTable(page, 2, 3);
    await clickVueTableCell(page, 0, 0);

    const firstLineMarker = page.getByRole('slider', { name: 'First line indent' });
    const leftIndentMarker = page.getByRole('slider', { name: 'Left indent' });
    const rightIndentMarker = page.getByRole('slider', { name: 'Right indent' });
    const firstCell = await visualTableCellContentRect(page, 0);

    await expect(firstLineMarker).toBeVisible();
    await expect(leftIndentMarker).toBeVisible();
    await expect(rightIndentMarker).toBeVisible();
    await expect(page.locator('.docx-ruler-table-boundary')).toHaveCount(3);

    expectClosePx(await markerCenterX(firstLineMarker), firstCell.left);
    expectClosePx(await markerCenterX(leftIndentMarker), firstCell.left);
    expectClosePx(await markerCenterX(rightIndentMarker), firstCell.right);

    await clickVueTableCell(page, 0, 1);
    const secondCell = await visualTableCellContentRect(page, 1);

    const firstLineX = await markerCenterX(firstLineMarker);
    expect(firstLineX).toBeGreaterThan(firstCell.right - 8);
    expectClosePx(firstLineX, secondCell.left);
    expectClosePx(await markerCenterX(leftIndentMarker), secondCell.left);
    expectClosePx(await markerCenterX(rightIndentMarker), secondCell.right);
  });

  test('typing in a new numbered item keeps the Vue caret on text', async ({ page }) => {
    await gotoVueEmpty(page);

    await page.locator('.layout-page-content').click();
    await page.keyboard.insertText('привет');
    await page.getByRole('button', { name: 'Numbered List' }).click();
    await page.keyboard.press('Enter');
    await page.keyboard.insertText('юпривет');

    const secondMarker = page.locator('.layout-page-content .layout-list-marker').nth(1);
    const typedRun = page
      .locator('.layout-page-content span[data-pm-start][data-pm-end]:not(.layout-list-marker)')
      .filter({ hasText: 'юпривет' })
      .last();
    const caret = page.locator('.vue-caret');

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

  test('clicking a list marker highlights the Vue marker group only after marker click', async ({
    page,
  }) => {
    await gotoVueEmpty(page);

    await page.locator('.layout-page-content').click();
    await page.keyboard.insertText('привет');
    await page.getByRole('button', { name: 'Numbered List' }).click();
    await page.keyboard.press('Enter');
    await page.keyboard.insertText('пока');
    await page.keyboard.press('Enter');
    await page.keyboard.insertText('привет');
    await page.keyboard.press('Enter');
    await page.keyboard.insertText('пока');

    const markers = page.locator('.layout-page-content .layout-list-marker');
    const selectedMarkers = page.locator('.layout-page-content .layout-list-marker-selected');
    await expect(markers).toHaveCount(4);
    await expect(selectedMarkers).toHaveCount(0);

    await markers.first().click();
    await expect(selectedMarkers).toHaveCount(4);

    const textRun = page
      .locator('.layout-page-content span[data-pm-start][data-pm-end]:not(.layout-list-marker)')
      .filter({ hasText: 'пока' })
      .first();
    await textRun.click();
    await expect(selectedMarkers).toHaveCount(0);
  });
});
