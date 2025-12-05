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

test.describe('Excel File Upload Tests', () => {
  test('should load simple single-sheet Excel file without errors', async ({ page }) => {
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
    await page.waitForLoadState('networkidle');

    // Upload the simple test file
    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'simple-single-sheet.xlsx'));

    // Wait for the graph to render (file name should appear in header)
    await expect(page.getByText('simple-single-sheet.xlsx')).toBeVisible({ timeout: 10000 });

    // Verify nodes and edges chips are visible (indicating data was parsed)
    await expect(page.getByText(/\d+ nodes/)).toBeVisible();
    await expect(page.getByText(/\d+ edges/)).toBeVisible();

    // Verify canvas is rendered (Force Graph uses canvas)
    await expect(page.locator('canvas')).toBeVisible();

    // Check for console errors (filtering known library issues)
    const filteredErrors = filterConsoleErrors(consoleErrors);
    expect(filteredErrors, `Console errors found: ${filteredErrors.join(', ')}`).toHaveLength(0);
  });

  test('should load multi-sheet Excel file with cross-references without errors', async ({ page }) => {
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
    await page.waitForLoadState('networkidle');

    // Upload the multi-sheet test file
    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'multi-sheet-cross-ref.xlsx'));

    // Wait for the graph to render
    await expect(page.getByText('multi-sheet-cross-ref.xlsx')).toBeVisible({ timeout: 10000 });

    // Verify nodes and edges are present
    await expect(page.getByText(/\d+ nodes/)).toBeVisible();
    await expect(page.getByText(/\d+ edges/)).toBeVisible();

    // Check for console errors (filtering known library issues)
    const filteredErrors = filterConsoleErrors(consoleErrors);
    expect(filteredErrors, `Console errors found: ${filteredErrors.join(', ')}`).toHaveLength(0);
  });

  test('should reset and load a new file', async ({ page }) => {
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

    // Upload first file
    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'simple-single-sheet.xlsx'));

    await expect(page.getByText('simple-single-sheet.xlsx')).toBeVisible({ timeout: 10000 });

    // Click reset button
    await page.getByRole('button', { name: 'Load new file' }).click();

    // Should be back to upload screen
    await expect(page.getByText('Drag & drop an Excel file here')).toBeVisible();

    // Upload second file
    const newFileInput = page.locator('input[type="file"]');
    await newFileInput.setInputFiles(path.join(FIXTURES_DIR, 'multi-sheet-cross-ref.xlsx'));

    await expect(page.getByText('multi-sheet-cross-ref.xlsx')).toBeVisible({ timeout: 10000 });

    // Check for console errors (filtering known library issues)
    const filteredErrors = filterConsoleErrors(consoleErrors);
    expect(filteredErrors, `Console errors found: ${filteredErrors.join(', ')}`).toHaveLength(0);
  });

  test('should display sheet filter dropdown', async ({ page }) => {
    await page.goto('/');

    const fileInput = page.locator('input[type="file"]');
    await fileInput.setInputFiles(path.join(FIXTURES_DIR, 'multi-sheet-cross-ref.xlsx'));

    await expect(page.getByText('multi-sheet-cross-ref.xlsx')).toBeVisible({ timeout: 10000 });

    // Verify sheet filter dropdown is visible by checking for "All Sheets" default value
    await expect(page.getByText('All Sheets')).toBeVisible();
  });
});
