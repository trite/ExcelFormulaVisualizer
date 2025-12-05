import { test, expect } from '@playwright/test';

test.describe('App Smoke Tests', () => {
  test('should load without console errors', async ({ page }) => {
    const consoleErrors: string[] = [];

    // Collect console errors
    page.on('console', (msg) => {
      if (msg.type() === 'error') {
        consoleErrors.push(msg.text());
      }
    });

    // Collect page errors (uncaught exceptions)
    page.on('pageerror', (err) => {
      consoleErrors.push(err.message);
    });

    // Navigate to the app
    await page.goto('/');

    // Wait for the page to be fully loaded
    await page.waitForLoadState('networkidle');

    // Check that the main heading is visible
    await expect(page.getByText('Visualize Excel Formula Relationships')).toBeVisible();

    // Check that the file upload area is present
    await expect(page.getByText('Drag & drop an Excel file here')).toBeVisible();

    // Assert no console errors
    expect(consoleErrors, `Console errors found: ${consoleErrors.join(', ')}`).toHaveLength(0);
  });

  test('should display the app header', async ({ page }) => {
    await page.goto('/');
    await expect(page.getByText('Excel Formula Visualizer')).toBeVisible();
  });

  test('should have working file upload button', async ({ page }) => {
    await page.goto('/');
    const browseButton = page.getByRole('button', { name: 'Browse Files' });
    await expect(browseButton).toBeVisible();
    await expect(browseButton).toBeEnabled();
  });
});
