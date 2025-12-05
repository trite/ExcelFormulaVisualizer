#!/usr/bin/env npx tsx
/**
 * Generate Excel test fixtures for named ranges testing
 *
 * Run with: npx tsx scripts/generate-test-fixtures.ts
 */

import * as XLSX from "xlsx";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const FIXTURES_DIR = path.join(__dirname, "..", "test-fixtures");

// Ensure fixtures directory exists
if (!fs.existsSync(FIXTURES_DIR)) {
  fs.mkdirSync(FIXTURES_DIR, { recursive: true });
}

/**
 * Helper to create a workbook with defined names
 */
function createWorkbookWithNames(
  sheets: {
    name: string;
    data: (string | number | { f: string })[][];
  }[],
  definedNames: {
    Name: string;
    Ref: string;
    Sheet?: number; // 0-indexed sheet scope, undefined for global
    Comment?: string;
  }[]
): XLSX.WorkBook {
  const workbook = XLSX.utils.book_new();

  // Add sheets
  for (const sheet of sheets) {
    // First create a worksheet with just the values
    const valuesOnly = sheet.data.map((row) =>
      row.map((cell) => {
        if (typeof cell === "object" && "f" in cell) {
          return undefined; // Will add formula separately
        }
        return cell;
      })
    );

    const worksheet = XLSX.utils.aoa_to_sheet(valuesOnly);

    // Track max row and column for updating range
    let maxRow = 0;
    let maxCol = 0;

    // Now add formulas directly to cells
    for (let r = 0; r < sheet.data.length; r++) {
      for (let c = 0; c < sheet.data[r].length; c++) {
        const cell = sheet.data[r][c];
        if (cell !== undefined) {
          maxRow = Math.max(maxRow, r);
          maxCol = Math.max(maxCol, c);
        }
        if (typeof cell === "object" && "f" in cell) {
          const cellAddr = XLSX.utils.encode_cell({ r, c });
          // Include both formula and a placeholder value (xlsx requires a value)
          worksheet[cellAddr] = { t: "n", v: 0, f: cell.f };
        }
      }
    }

    // Update the worksheet range to include all cells
    worksheet["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: maxRow, c: maxCol },
    });

    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
  }

  // Add defined names
  if (definedNames.length > 0) {
    if (!workbook.Workbook) {
      workbook.Workbook = {};
    }
    workbook.Workbook.Names = definedNames;
  }

  return workbook;
}

function saveWorkbook(workbook: XLSX.WorkBook, filename: string): void {
  const filepath = path.join(FIXTURES_DIR, filename);
  XLSX.writeFile(workbook, filepath);
  console.log(`Created: ${filepath}`);
}

// =============================================================================
// Test Fixture 1: Global named cell
// =============================================================================
function createGlobalNamedCell(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Sheet1",
        data: [
          [100, { f: "mass*2" }], // A1=100 (named "mass"), B1=mass*2
          [{ f: "mass+10" }], // A2=mass+10
        ],
      },
    ],
    [{ Name: "mass", Ref: "Sheet1!$A$1" }]
  );

  saveWorkbook(workbook, "global-named-cell.xlsx");
}

// =============================================================================
// Test Fixture 2: Global named range
// =============================================================================
function createGlobalNamedRange(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Sheet1",
        data: [
          [10], // A1
          [20], // A2
          [30], // A3
          [{ f: "SUM(myRange)" }], // A4=SUM(myRange)
        ],
      },
    ],
    [{ Name: "myRange", Ref: "Sheet1!$A$1:$A$3" }]
  );

  saveWorkbook(workbook, "global-named-range.xlsx");
}

// =============================================================================
// Test Fixture 3: Sheet-scoped name (non-ambiguous)
// =============================================================================
function createSheetScopedName(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Sheet1",
        data: [
          [42, { f: "localData+1" }], // A1=42 (named "localData" scoped to Sheet1), B1=localData+1
        ],
      },
      {
        name: "Sheet2",
        data: [
          [99], // A1=99 (no name)
        ],
      },
    ],
    [{ Name: "localData", Ref: "Sheet1!$A$1", Sheet: 0 }]
  );

  saveWorkbook(workbook, "sheet-scoped-name.xlsx");
}

// =============================================================================
// Test Fixture 4: Ambiguous names (same name in multiple scopes)
// =============================================================================
function createAmbiguousNames(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Sheet1",
        data: [
          [100, { f: "data*2" }], // A1=100 (named "data" scoped to Sheet1), B1=data*2
        ],
      },
      {
        name: "Sheet2",
        data: [
          [200, { f: "data*3" }], // A1=200 (named "data" scoped to Sheet2), B1=data*3
        ],
      },
    ],
    [
      { Name: "data", Ref: "Sheet1!$A$1", Sheet: 0 },
      { Name: "data", Ref: "Sheet2!$A$1", Sheet: 1 },
    ]
  );

  saveWorkbook(workbook, "ambiguous-names.xlsx");
}

// =============================================================================
// Test Fixture 5: Global and scoped name with same name
// =============================================================================
function createGlobalAndScopedSameName(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Sheet1",
        data: [
          [50, { f: "value+1" }], // A1=50 (scoped "value"), B1=value+1 (uses scoped)
        ],
      },
      {
        name: "Sheet2",
        data: [
          [999, { f: "value+2" }], // A1=999 (global "value" target), B1=value+2 (uses global)
        ],
      },
    ],
    [
      { Name: "value", Ref: "Sheet2!$A$1" }, // Global
      { Name: "value", Ref: "Sheet1!$A$1", Sheet: 0 }, // Scoped to Sheet1
    ]
  );

  saveWorkbook(workbook, "global-and-scoped-same-name.xlsx");
}

// =============================================================================
// Test Fixture 6: Names with special characters (underscores, dots)
// =============================================================================
function createNamesWithSpecialChars(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Sheet1",
        data: [
          [1, 2, 3], // A1, B1, C1
          [{ f: "my_value+tax.rate+data_2024" }], // A2 references all three names
        ],
      },
    ],
    [
      { Name: "my_value", Ref: "Sheet1!$A$1" },
      { Name: "tax.rate", Ref: "Sheet1!$B$1" },
      { Name: "data_2024", Ref: "Sheet1!$C$1" },
    ]
  );

  saveWorkbook(workbook, "names-with-special-chars.xlsx");
}

// =============================================================================
// Test Fixture 7: Quoted sheet name with spaces
// =============================================================================
function createQuotedSheetName(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "My Data Sheet",
        data: [
          [500], // A1=500
        ],
      },
      {
        name: "Calculations",
        data: [
          [{ f: "revenue*2" }], // A1=revenue*2
        ],
      },
    ],
    [{ Name: "revenue", Ref: "'My Data Sheet'!$A$1" }]
  );

  saveWorkbook(workbook, "quoted-sheet-name.xlsx");
}

// =============================================================================
// Test Fixture 8: Multiple names pointing to same cell
// =============================================================================
function createMultipleNamesToSameCell(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Sheet1",
        data: [
          [3.14159, { f: "pi*radius" }, { f: "circumference" }], // A1=pi, B1=pi*radius, C1=circumference
          [5], // A2=5 (radius)
        ],
      },
    ],
    [
      { Name: "pi", Ref: "Sheet1!$A$1" },
      { Name: "circumference", Ref: "Sheet1!$A$1" }, // Same cell, different name
      { Name: "radius", Ref: "Sheet1!$A$2" },
    ]
  );

  saveWorkbook(workbook, "multiple-names-same-cell.xlsx");
}

// =============================================================================
// Test Fixture 9: Cross-sheet formula with named reference
// =============================================================================
function createCrossSheetNamedRef(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Data",
        data: [
          [1000], // A1=1000 (named "baseValue")
        ],
      },
      {
        name: "Calculations",
        data: [
          [{ f: "baseValue*1.1" }], // A1=baseValue*1.1
          [{ f: "Data!A1*1.2" }], // A2=Data!A1*1.2 (direct ref, not named)
        ],
      },
    ],
    [{ Name: "baseValue", Ref: "Data!$A$1" }]
  );

  saveWorkbook(workbook, "cross-sheet-named-ref.xlsx");
}

// =============================================================================
// Test Fixture 10: Complex scenario with mixed references
// =============================================================================
function createComplexScenario(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Inputs",
        data: [
          [100, 0.05, 12], // A1=principal, B1=rate, C1=months
        ],
      },
      {
        name: "Calculations",
        data: [
          [{ f: "principal*(1+rate)^months" }], // A1=compound interest formula
          [{ f: "principal+Inputs!B1*100" }], // A2=mixed named and direct ref
          [{ f: "localMultiplier*principal" }], // A3=uses sheet-scoped name
          [2], // A4=2 (localMultiplier, scoped to Calculations)
        ],
      },
      {
        name: "Summary",
        data: [
          [{ f: "principal" }], // A1=references global name from another sheet
          [{ f: "Calculations!A1" }], // A2=direct cell reference
        ],
      },
    ],
    [
      { Name: "principal", Ref: "Inputs!$A$1" },
      { Name: "rate", Ref: "Inputs!$B$1" },
      { Name: "months", Ref: "Inputs!$C$1" },
      { Name: "localMultiplier", Ref: "Calculations!$A$4", Sheet: 1 }, // Scoped to Calculations
    ]
  );

  saveWorkbook(workbook, "complex-scenario.xlsx");
}

// =============================================================================
// Test Fixture 11: Name that could look like a cell reference
// =============================================================================
function createNameLikeCellRef(): void {
  // Excel actually prevents names like "A1" but allows things like "A1_data"
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Sheet1",
        data: [
          [10, 20], // A1=10, B1=20
          [{ f: "A1_data+B1_data" }], // A2 references the names
        ],
      },
    ],
    [
      { Name: "A1_data", Ref: "Sheet1!$A$1" },
      { Name: "B1_data", Ref: "Sheet1!$B$1" },
    ]
  );

  saveWorkbook(workbook, "name-like-cell-ref.xlsx");
}

// =============================================================================
// Test Fixture 12: No defined names (control case)
// =============================================================================
function createNoDefinedNames(): void {
  const workbook = createWorkbookWithNames(
    [
      {
        name: "Sheet1",
        data: [
          [10, { f: "A1*2" }], // A1=10, B1=A1*2
          [{ f: "A1+B1" }], // A2=A1+B1
        ],
      },
    ],
    [] // No defined names
  );

  saveWorkbook(workbook, "no-defined-names.xlsx");
}

// =============================================================================
// Main
// =============================================================================
function main(): void {
  console.log("Generating test fixtures...\n");

  createGlobalNamedCell();
  createGlobalNamedRange();
  createSheetScopedName();
  createAmbiguousNames();
  createGlobalAndScopedSameName();
  createNamesWithSpecialChars();
  createQuotedSheetName();
  createMultipleNamesToSameCell();
  createCrossSheetNamedRef();
  createComplexScenario();
  createNameLikeCellRef();
  createNoDefinedNames();

  console.log("\nDone! Test fixtures created in:", FIXTURES_DIR);
}

main();
