import { test, expect, Page } from '@playwright/test';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const FIXTURES_DIR = path.join(__dirname, 'fixtures');

// Known benign errors from third-party libraries that don't affect functionality
const KNOWN_LIBRARY_ERRORS = [
  "Cannot read properties of null (reading 'notify')", // react-force-graph-2d d3 cleanup
];

function isKnownLibraryError(error: string): boolean {
  return KNOWN_LIBRARY_ERRORS.some(known => error.includes(known));
}

function filterConsoleErrors(errors: string[]): string[] {
  return errors.filter(err => !isKnownLibraryError(err));
}

/**
 * Try to click on a node by attempting clicks at multiple positions on the canvas.
 * Force graphs have dynamic node positions, so we try a grid of positions.
 * Returns true if a node was selected (detected by node details panel showing cell info).
 */
async function trySelectNode(page: Page, maxAttempts = 16): Promise<boolean> {
  const canvas = page.locator('canvas').first();
  const box = await canvas.boundingBox();
  if (!box) return false;

  // Try a 4x4 grid of positions within the canvas (more positions for better coverage)
  const positions = [];
  for (let row = 0; row < 4; row++) {
    for (let col = 0; col < 4; col++) {
      positions.push({
        x: 0.2 + (col * 0.2),
        y: 0.2 + (row * 0.2),
      });
    }
  }

  for (let i = 0; i < Math.min(maxAttempts, positions.length); i++) {
    const pos = positions[i];
    await page.mouse.click(box.x + box.width * pos.x, box.y + box.height * pos.y);
    await page.waitForTimeout(400);

    // Check if we selected a node by looking for various indicators in the details panel
    // A selected node will show: "Add Name" button, "Add Note" button, cell address, or "Name" section
    const indicators = [
      page.getByRole('button', { name: 'Add Name' }),
      page.getByRole('button', { name: 'Add Note' }),
      page.getByText('Cell Address'),
      page.locator('text=/Sheet\\d?!\\w+\\d+/'), // Matches patterns like Sheet1!A1
    ];

    for (const indicator of indicators) {
      const isVisible = await indicator.first().isVisible().catch(() => false);
      if (isVisible) {
        return true;
      }
    }
  }

  return false;
}

test.describe('Cell Metadata - Names', () => {
  test.beforeEach(async ({ page }) => {
    // Clear localStorage before each test
    await page.goto('/');
    await page.evaluate(() => localStorage.clear());
  });

  test('should show "Add Name" button when node is selected', async ({ page }) => {
    await page.goto('/');

    // Upload test file
    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    // Wait for graph to render and stabilize
    await page.waitForTimeout(2000);

    // Try to select a node
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    // Should see Add Name button in node details
    await expect(page.getByRole('button', { name: 'Add Name' })).toBeVisible({ timeout: 5000 });
  });

  test('should add a name to a cell and display it', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    // Click Add Name button
    await page.getByRole('button', { name: 'Add Name' }).click();

    // Enter a name
    const nameInput = page.locator('input[placeholder="Enter a name for this cell"]');
    await nameInput.fill('TestCellName');

    // Click Save
    await page.getByRole('button', { name: 'Save' }).click();

    // Should see the name displayed as a chip
    await expect(page.getByText('TestCellName')).toBeVisible();
  });

  test('should persist name after page reload', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    // Add a name
    await page.getByRole('button', { name: 'Add Name' }).click();
    const nameInput = page.locator('input[placeholder="Enter a name for this cell"]');
    await nameInput.fill('PersistentName');
    await page.getByRole('button', { name: 'Save' }).click();

    await expect(page.getByText('PersistentName')).toBeVisible();

    // Reload the page
    await page.reload();

    // Re-upload the same file
    const newFileInput = page.locator('input[type="file"]');
    await newFileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node again - note: may not be the same node due to force layout
    // But we should see "annotations" count in toolbar if metadata persisted
    await expect(page.getByText(/\d+ annotations/)).toBeVisible({ timeout: 5000 });
  });

  test('should edit an existing name', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    // Add initial name
    await page.getByRole('button', { name: 'Add Name' }).click();
    await page.locator('input[placeholder="Enter a name for this cell"]').fill('OriginalName');
    await page.getByRole('button', { name: 'Save' }).click();

    // Click edit button (the pencil icon next to the name)
    await page.getByRole('button', { name: 'Edit name' }).click();

    // Change the name
    const editInput = page.locator('input[placeholder="Enter a name for this cell"]');
    await editInput.clear();
    await editInput.fill('UpdatedName');
    await page.getByRole('button', { name: 'Save' }).click();

    // Should see updated name
    await expect(page.getByText('UpdatedName')).toBeVisible();
    await expect(page.getByText('OriginalName')).not.toBeVisible();
  });

  test('should delete a name', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    // Add a name
    await page.getByRole('button', { name: 'Add Name' }).click();
    await page.locator('input[placeholder="Enter a name for this cell"]').fill('ToDelete');
    await page.getByRole('button', { name: 'Save' }).click();

    await expect(page.getByText('ToDelete')).toBeVisible();

    // The name chip should have a delete button (X icon)
    // MUI Chip with onDelete shows a cancel icon
    const chip = page.locator('.MuiChip-root').filter({ hasText: 'ToDelete' });
    await chip.locator('[data-testid="CancelIcon"]').click();

    // Name should be gone, Add Name button should be back
    await expect(page.getByText('ToDelete')).not.toBeVisible();
    await expect(page.getByRole('button', { name: 'Add Name' })).toBeVisible();
  });
});

test.describe('Cell Metadata - Notes', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await page.evaluate(() => localStorage.clear());
  });

  test('should add a note to a cell', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    // Click Add Note button
    await page.getByRole('button', { name: 'Add Note' }).click();

    // Enter a note
    const noteInput = page.locator('textarea[placeholder="Enter a note"]');
    await noteInput.fill('This is a test note');

    // Click Save
    await page.getByRole('button', { name: 'Save' }).click();

    // Should see the note displayed
    await expect(page.getByText('This is a test note')).toBeVisible();
  });

  test('should allow multiple notes on same cell', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    // Add first note
    await page.getByRole('button', { name: 'Add Note' }).click();
    await page.locator('textarea[placeholder="Enter a note"]').fill('First note');
    await page.getByRole('button', { name: 'Save' }).click();

    // Add second note
    await page.getByRole('button', { name: 'Add Note' }).click();
    await page.locator('textarea[placeholder="Enter a note"]').fill('Second note');
    await page.getByRole('button', { name: 'Save' }).click();

    // Add third note
    await page.getByRole('button', { name: 'Add Note' }).click();
    await page.locator('textarea[placeholder="Enter a note"]').fill('Third note');
    await page.getByRole('button', { name: 'Save' }).click();

    // All notes should be visible
    await expect(page.getByText('First note')).toBeVisible();
    await expect(page.getByText('Second note')).toBeVisible();
    await expect(page.getByText('Third note')).toBeVisible();

    // Notes count should show 3
    await expect(page.getByText('Notes (3)')).toBeVisible();
  });

  test('should persist notes after page reload', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    // Add a note
    await page.getByRole('button', { name: 'Add Note' }).click();
    await page.locator('textarea[placeholder="Enter a note"]').fill('Persistent note');
    await page.getByRole('button', { name: 'Save' }).click();

    await expect(page.getByText('Persistent note')).toBeVisible();

    // Reload the page
    await page.reload();

    // Re-upload the same file
    const newFileInput = page.locator('input[type="file"]');
    await newFileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // The note should have persisted - check annotations count in toolbar
    await expect(page.getByText(/\d+ annotations/)).toBeVisible({ timeout: 5000 });
  });
});

test.describe('Cell Metadata - Clear All', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await page.evaluate(() => localStorage.clear());
  });

  test('should show confirmation dialog when clearing metadata', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node and add some metadata
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    await page.getByRole('button', { name: 'Add Name' }).click();
    await page.locator('input[placeholder="Enter a name for this cell"]').fill('ToBeCleared');
    await page.getByRole('button', { name: 'Save' }).click();

    // Click clear metadata button in toolbar
    await page.getByRole('button', { name: 'Clear all metadata' }).click();

    // Should show confirmation dialog
    await expect(page.getByText('Clear All Metadata?')).toBeVisible();
    await expect(page.getByRole('button', { name: 'Clear All' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Cancel' })).toBeVisible();
  });

  test('should clear all metadata when confirmed', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node and add metadata
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    await page.getByRole('button', { name: 'Add Name' }).click();
    await page.locator('input[placeholder="Enter a name for this cell"]').fill('ToClear');
    await page.getByRole('button', { name: 'Save' }).click();

    // Verify annotation count appears in toolbar
    await expect(page.getByText(/\d+ annotations/)).toBeVisible();

    // Clear metadata
    await page.getByRole('button', { name: 'Clear all metadata' }).click();
    await page.getByRole('button', { name: 'Clear All' }).click();

    // Wait for dialog to close
    await page.waitForTimeout(500);

    // The annotations chip should be gone from toolbar
    await expect(page.getByText(/\d+ annotations/)).not.toBeVisible();
  });

  test('should not clear metadata when cancelled', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await page.waitForTimeout(2000);

    // Select a node and add metadata
    const selected = await trySelectNode(page);
    expect(selected, 'Should be able to select a node').toBe(true);

    await page.getByRole('button', { name: 'Add Name' }).click();
    await page.locator('input[placeholder="Enter a name for this cell"]').fill('ShouldRemain');
    await page.getByRole('button', { name: 'Save' }).click();

    await expect(page.getByText('ShouldRemain')).toBeVisible();

    // Click clear but cancel
    await page.getByRole('button', { name: 'Clear all metadata' }).click();
    await page.getByRole('button', { name: 'Cancel' }).click();

    // Name should still be there
    await expect(page.getByText('ShouldRemain')).toBeVisible();
  });
});

test.describe('Cell Metadata - Import from Excel', () => {
  test('should load metadata from Excel file', async ({ page }) => {
    const consoleErrors: string[] = [];

    page.on('console', (msg) => {
      if (msg.type() === 'error') {
        consoleErrors.push(msg.text());
      }
    });

    await page.goto('/');
    await page.evaluate(() => localStorage.clear());

    // Upload file with pre-existing metadata
    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'with-metadata.xlsx'));

    await page.waitForTimeout(1500);

    // Select the first node (Sheet1!A1)
    const canvas = page.locator('canvas').first();
    const box = await canvas.boundingBox();
    if (box) {
      await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
    }

    await page.waitForTimeout(500);

    // Should see metadata from the Excel file
    // Note: The exact node clicked may vary, but the file has metadata
    // Check that annotations count shows in toolbar
    await expect(page.getByText(/\d+ annotations/)).toBeVisible({ timeout: 5000 });

    // Check console errors
    const filteredErrors = filterConsoleErrors(consoleErrors);
    expect(filteredErrors, `Console errors found: ${filteredErrors.join(', ')}`).toHaveLength(0);
  });
});

test.describe('Graph Edge Verification', () => {
  test('should have correct number of nodes and edges for edge-test fixture', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'edge-test.xlsx'));

    // Wait for graph to load
    await expect(page.getByText('edge-test.xlsx')).toBeVisible({ timeout: 10000 });

    // Check node count (should be 6: A1, A2, A3, B1 from Sheet1, A1, A2 from Sheet2)
    await expect(page.getByText('6 nodes')).toBeVisible();

    // Check edge count (should be 6 based on the fixture structure)
    await expect(page.getByText('6 edges')).toBeVisible();
  });

  test('should have correct counts for metadata-test fixture', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'metadata-test.xlsx'));

    await expect(page.getByText('metadata-test.xlsx')).toBeVisible({ timeout: 10000 });

    // metadata-test.xlsx has:
    // A1: 100, A2: 200, B1: =A1+A2, B2: =A1*2
    // Nodes: A1, A2 (referenced), B1, B2 (formulas) = 4 nodes
    // Edges: A1->B1, A2->B1, A1->B2 = 3 edges
    await expect(page.getByText('4 nodes')).toBeVisible();
    await expect(page.getByText('3 edges')).toBeVisible();
  });
});
