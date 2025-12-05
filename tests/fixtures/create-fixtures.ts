/**
 * Script to create test fixture Excel files with known structures
 * Run with: npx tsx tests/fixtures/create-fixtures.ts
 */
import * as XLSX from 'xlsx';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * metadata-test.xlsx
 * Simple structure for metadata testing
 * Sheet1:
 *   A1: 100 (value only - will be referenced)
 *   A2: 200 (value only - will be referenced)
 *   B1: =A1+A2 (formula referencing A1, A2)
 *   B2: =A1*2 (formula referencing A1 only)
 */
function createMetadataTestFile() {
  const wb = XLSX.utils.book_new();

  const ws = XLSX.utils.aoa_to_sheet([
    [100, { f: 'A1+A2' }],
    [200, { f: 'A1*2' }],
  ]);

  // Set cell types correctly
  ws['A1'] = { t: 'n', v: 100 };
  ws['A2'] = { t: 'n', v: 200 };
  ws['B1'] = { t: 'n', v: 300, f: 'A1+A2' };
  ws['B2'] = { t: 'n', v: 200, f: 'A1*2' };
  ws['!ref'] = 'A1:B2';

  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  const filePath = path.join(__dirname, 'metadata-test.xlsx');
  XLSX.writeFile(wb, filePath);
  console.log(`Created: ${filePath}`);
}

/**
 * edge-test.xlsx
 * Structure for testing edge generation
 * Sheet1:
 *   A1: 10
 *   A2: 20
 *   A3: =A1+A2
 *   B1: =A3*2
 * Sheet2:
 *   A1: =Sheet1!A1+Sheet1!A2
 *   A2: =Sheet1!B1
 *
 * Expected edges:
 *   Sheet1!A1 → Sheet1!A3
 *   Sheet1!A2 → Sheet1!A3
 *   Sheet1!A3 → Sheet1!B1
 *   Sheet1!A1 → Sheet2!A1
 *   Sheet1!A2 → Sheet2!A1
 *   Sheet1!B1 → Sheet2!A2
 */
function createEdgeTestFile() {
  const wb = XLSX.utils.book_new();

  // Sheet1
  const ws1: XLSX.WorkSheet = {
    'A1': { t: 'n', v: 10 },
    'A2': { t: 'n', v: 20 },
    'A3': { t: 'n', v: 30, f: 'A1+A2' },
    'B1': { t: 'n', v: 60, f: 'A3*2' },
    '!ref': 'A1:B3',
  };

  XLSX.utils.book_append_sheet(wb, ws1, 'Sheet1');

  // Sheet2
  const ws2: XLSX.WorkSheet = {
    'A1': { t: 'n', v: 30, f: 'Sheet1!A1+Sheet1!A2' },
    'A2': { t: 'n', v: 60, f: 'Sheet1!B1' },
    '!ref': 'A1:A2',
  };

  XLSX.utils.book_append_sheet(wb, ws2, 'Sheet2');

  const filePath = path.join(__dirname, 'edge-test.xlsx');
  XLSX.writeFile(wb, filePath);
  console.log(`Created: ${filePath}`);
}

/**
 * range-name-test.xlsx
 * Structure for testing range handling in metadata
 * Sheet1:
 *   A1:A5: Values 1-5
 *   B1: =SUM(A1:A5)
 */
function createRangeNameTestFile() {
  const wb = XLSX.utils.book_new();

  const ws: XLSX.WorkSheet = {
    'A1': { t: 'n', v: 1 },
    'A2': { t: 'n', v: 2 },
    'A3': { t: 'n', v: 3 },
    'A4': { t: 'n', v: 4 },
    'A5': { t: 'n', v: 5 },
    'B1': { t: 'n', v: 15, f: 'SUM(A1:A5)' },
    '!ref': 'A1:B5',
  };

  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  const filePath = path.join(__dirname, 'range-name-test.xlsx');
  XLSX.writeFile(wb, filePath);
  console.log(`Created: ${filePath}`);
}

/**
 * with-metadata.xlsx
 * File with pre-existing metadata sheet for testing import
 * Sheet1:
 *   A1: 100
 *   A2: =A1*2
 * ExcelFormulaVisualizerMetadata (hidden):
 *   Name Keys | Name Values | Note Keys | Note Values
 *   Sheet1!A1 | MyValue     | Sheet1!A1 | This is a test note
 *   Sheet1!A2 | DoubleValue | Sheet1!A2 | Formula cell note
 */
function createWithMetadataFile() {
  const wb = XLSX.utils.book_new();

  // Sheet1
  const ws1: XLSX.WorkSheet = {
    'A1': { t: 'n', v: 100 },
    'A2': { t: 'n', v: 200, f: 'A1*2' },
    '!ref': 'A1:A2',
  };

  XLSX.utils.book_append_sheet(wb, ws1, 'Sheet1');

  // Metadata sheet
  const metadataWs = XLSX.utils.aoa_to_sheet([
    ['Name Keys', 'Name Values', 'Note Keys', 'Note Values'],
    ['Sheet1!A1', 'MyValue', 'Sheet1!A1', 'This is a test note'],
    ['Sheet1!A2', 'DoubleValue', 'Sheet1!A2', 'Formula cell note'],
  ]);

  metadataWs['!cols'] = [
    { wch: 30 },
    { wch: 20 },
    { wch: 30 },
    { wch: 40 },
  ];

  XLSX.utils.book_append_sheet(wb, metadataWs, 'ExcelFormulaVisualizerMetadata');

  // Hide the metadata sheet
  if (!wb.Workbook) {
    wb.Workbook = { Sheets: [] };
  }
  if (!wb.Workbook.Sheets) {
    wb.Workbook.Sheets = [];
  }
  while (wb.Workbook.Sheets.length < wb.SheetNames.length) {
    wb.Workbook.Sheets.push({});
  }
  const metadataSheetIdx = wb.SheetNames.indexOf('ExcelFormulaVisualizerMetadata');
  wb.Workbook.Sheets[metadataSheetIdx] = { Hidden: 1 };

  const filePath = path.join(__dirname, 'with-metadata.xlsx');
  XLSX.writeFile(wb, filePath);
  console.log(`Created: ${filePath}`);
}

// Create all fixtures
console.log('Creating test fixtures...');
createMetadataTestFile();
createEdgeTestFile();
createRangeNameTestFile();
createWithMetadataFile();
console.log('Done!');
