import { test, expect, type Page } from '@playwright/test';

const VUE_EMPTY_URL = 'http://localhost:5174/?e2e=1&empty=1';

async function gotoVueEmpty(page: Page): Promise<void> {
  await page.goto(VUE_EMPTY_URL);
  await page.locator('.docx-editor-vue').waitFor();
  await page.locator('.paged-editor__pages').waitFor();
  await page.waitForSelector('[data-page-number]');
}

async function selectAllVueDocument(page: Page): Promise<void> {
  await page.evaluate(() => {
    const contentArea = document.querySelector('.ProseMirror');
    if (!contentArea) return;

    const walker = document.createTreeWalker(contentArea, NodeFilter.SHOW_TEXT, null);
    let firstTextNode: Text | null = null;
    let lastTextNode: Text | null = null;

    while (walker.nextNode()) {
      const text = walker.currentNode.textContent || '';
      if (text.length > 0) {
        firstTextNode ??= walker.currentNode as Text;
        lastTextNode = walker.currentNode as Text;
      }
    }

    if (!firstTextNode || !lastTextNode) return;
    const selection = window.getSelection();
    if (!selection) return;

    selection.removeAllRanges();
    const range = document.createRange();
    range.setStart(firstTextNode, 0);
    range.setEnd(lastTextNode, lastTextNode.textContent?.length || 0);
    selection.addRange(range);
  });
}

async function setVueParagraphSpacingPoints(
  page: Page,
  side: 'before' | 'after',
  points: number
): Promise<void> {
  await page.locator('.basic-toolbar [title="Line spacing"]').click();
  const inputName =
    side === 'before' ? 'Spacing before paragraph (pt)' : 'Spacing after paragraph (pt)';
  const input = page.getByRole('spinbutton', { name: inputName, exact: true });
  await input.fill(String(points));
  await input.press('Enter');
}

test.describe('Vue line spacing', () => {
  test('paragraph spacing accepts point values', async ({ page }) => {
    await gotoVueEmpty(page);

    await page.locator('.layout-page-content').click();
    await page.keyboard.insertText('First paragraph');
    await page.keyboard.press('Enter');
    await page.keyboard.insertText('Second paragraph');
    await selectAllVueDocument(page);

    await setVueParagraphSpacingPoints(page, 'before', 6);
    await setVueParagraphSpacingPoints(page, 'after', 18.5);

    const attrsByParagraph = await page.evaluate(() => [
      window.__DOCX_EDITOR_E2E__?.getParagraphAttrs(0),
      window.__DOCX_EDITOR_E2E__?.getParagraphAttrs(1),
    ]);
    for (const attrs of attrsByParagraph) {
      expect(attrs?.spaceBefore).toBe(120);
      expect(attrs?.spaceAfter).toBe(370);
      expect(attrs?.spacingExplicit).toMatchObject({ before: true, after: true });
    }
  });
});
