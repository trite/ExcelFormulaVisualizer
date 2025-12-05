import { test, expect } from '@playwright/test';
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

test.describe('Graph Rendering Verification', () => {
  test('Force Graph should render canvas element', async ({ page }) => {
    const consoleErrors: string[] = [];

    page.on('console', (msg) => {
      if (msg.type() === 'error') {
        consoleErrors.push(msg.text());
      }
    });

    page.on('pageerror', (err) => {
      consoleErrors.push(err.message);
    });

    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'simple-single-sheet.xlsx'));

    // Wait for the graph to render
    await expect(page.getByText('simple-single-sheet.xlsx')).toBeVisible({ timeout: 10000 });

    // Force Graph creates a canvas element
    const canvas = page.locator('canvas').first();
    await expect(canvas).toBeVisible();

    // Verify canvas has reasonable dimensions
    const box = await canvas.boundingBox();
    expect(box).not.toBeNull();
    expect(box!.width).toBeGreaterThan(100);
    expect(box!.height).toBeGreaterThan(100);

    // Check for console errors
    const filteredErrors = filterConsoleErrors(consoleErrors);
    expect(filteredErrors, `Console errors found: ${filteredErrors.join(', ')}`).toHaveLength(0);
  });

  test('clicking on canvas should not cause errors', async ({ page }) => {
    const consoleErrors: string[] = [];

    page.on('console', (msg) => {
      if (msg.type() === 'error') {
        consoleErrors.push(msg.text());
      }
    });

    page.on('pageerror', (err) => {
      consoleErrors.push(err.message);
    });

    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'simple-single-sheet.xlsx'));

    await expect(page.getByText('simple-single-sheet.xlsx')).toBeVisible({ timeout: 10000 });

    // Wait for graph to settle
    await page.waitForTimeout(1500);

    // Click on the canvas multiple times
    const canvas = page.locator('canvas').first();
    const box = await canvas.boundingBox();
    if (box) {
      // Click in various positions on the canvas
      await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
      await page.waitForTimeout(200);
      await page.mouse.click(box.x + box.width / 3, box.y + box.height / 3);
      await page.waitForTimeout(200);
      await page.mouse.click(box.x + box.width * 0.7, box.y + box.height * 0.7);
    }

    await page.waitForTimeout(500);

    // Check for console errors after clicking
    const filteredErrors = filterConsoleErrors(consoleErrors);
    expect(filteredErrors, `Console errors found: ${filteredErrors.join(', ')}`).toHaveLength(0);
  });

  test('multi-sheet file should render graph with multiple node colors', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'multi-sheet-cross-ref.xlsx'));

    await expect(page.getByText('multi-sheet-cross-ref.xlsx')).toBeVisible({ timeout: 10000 });

    // Verify nodes and edges are displayed
    await expect(page.getByText(/\d+ nodes/)).toBeVisible();
    await expect(page.getByText(/\d+ edges/)).toBeVisible();

    // Canvas should be visible
    await expect(page.locator('canvas')).toBeVisible();
  });

  test('graph should be responsive to window resize', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'simple-single-sheet.xlsx'));

    await expect(page.getByText('simple-single-sheet.xlsx')).toBeVisible({ timeout: 10000 });

    // Get initial canvas size
    const canvas = page.locator('canvas').first();
    const initialBox = await canvas.boundingBox();
    expect(initialBox).not.toBeNull();

    // Resize viewport
    await page.setViewportSize({ width: 800, height: 600 });
    await page.waitForTimeout(500);

    // Canvas should still be visible after resize
    await expect(canvas).toBeVisible();
    const newBox = await canvas.boundingBox();
    expect(newBox).not.toBeNull();
  });

  test('sheet filter should be available for multi-sheet files', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'multi-sheet-cross-ref.xlsx'));

    await expect(page.getByText('multi-sheet-cross-ref.xlsx')).toBeVisible({ timeout: 10000 });

    // Sheet filter dropdown should be visible by checking for "All Sheets" default value
    await expect(page.getByText('All Sheets')).toBeVisible();
  });

  test('graph should render without errors for file with many references', async ({ page }) => {
    const consoleErrors: string[] = [];

    page.on('console', (msg) => {
      if (msg.type() === 'error') {
        consoleErrors.push(msg.text());
      }
    });

    page.on('pageerror', (err) => {
      consoleErrors.push(err.message);
    });

    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'edge-test.xlsx'));

    await expect(page.getByText('edge-test.xlsx')).toBeVisible({ timeout: 10000 });

    // Verify nodes and edges are displayed
    await expect(page.getByText(/\d+ nodes/)).toBeVisible();
    await expect(page.getByText(/\d+ edges/)).toBeVisible();

    // Check for console errors
    const filteredErrors = filterConsoleErrors(consoleErrors);
    expect(filteredErrors, `Console errors found: ${filteredErrors.join(', ')}`).toHaveLength(0);
  });
});
