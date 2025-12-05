/**
 * Integration tests for Excel named ranges functionality
 *
 * These tests use actual Excel fixture files to verify the full pipeline:
 * 1. Excel file parsing extracts defined names
 * 2. Formula parser resolves named references to cell addresses
 * 3. Graph generation creates correct edges for named references
 * 4. Display names show correct format (with/without sheet qualifier)
 */

import { describe, it, expect, beforeAll } from "vitest";
import * as fs from "fs";
import * as path from "path";
import * as XLSX from "xlsx";
import { createNameResolver } from "../nameResolver";
import { extractCellReferences } from "../formulaParser";
import { workbookToGraphData } from "../excelParser";
import type { DefinedNameInfo, WorkbookData, CellInfo, SheetData } from "../../types";

// Path to test fixtures
const FIXTURES_DIR = path.join(__dirname, "../../../test-fixtures");

/**
 * Helper to load an Excel file and extract defined names
 */
function loadWorkbookWithNames(filename: string): {
  workbook: XLSX.WorkBook;
  definedNames: DefinedNameInfo[];
  sheetNames: string[];
} {
  const filepath = path.join(FIXTURES_DIR, filename);
  const workbook = XLSX.read(fs.readFileSync(filepath), {
    cellFormula: true,
    cellNF: true,
  });

  const definedNames: DefinedNameInfo[] = [];
  const names = workbook.Workbook?.Names;
  if (names && Array.isArray(names)) {
    for (const name of names) {
      if (!name.Name || name.Name.startsWith("_xlnm.")) continue;
      if (!name.Ref || name.Ref.startsWith("#")) continue;
      definedNames.push({
        name: name.Name,
        ref: name.Ref,
        sheetScope: name.Sheet,
        comment: name.Comment,
      });
    }
  }

  return {
    workbook,
    definedNames,
    sheetNames: workbook.SheetNames,
  };
}

/**
 * Helper to parse a workbook into WorkbookData format
 */
function parseWorkbookToData(
  workbook: XLSX.WorkBook,
  definedNames: DefinedNameInfo[],
  filename: string
): WorkbookData {
  const sheetNames = workbook.SheetNames;
  const nameResolver = createNameResolver(definedNames, sheetNames);

  const sheets: SheetData[] = [];
  const allCells = new Map<string, CellInfo>();

  for (let sheetIndex = 0; sheetIndex < sheetNames.length; sheetIndex++) {
    const sheetName = sheetNames[sheetIndex];
    const worksheet = workbook.Sheets[sheetName];
    const sheetCells = new Map<string, CellInfo>();

    const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1");

    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[cellAddress];

        if (cell) {
          const fullAddress = `${sheetName}!${cellAddress}`;
          const formula = cell.f ? `=${cell.f}` : undefined;
          const references = formula
            ? extractCellReferences(formula, sheetName, {
                nameResolver,
                currentSheetIndex: sheetIndex,
              })
            : [];

          const excelName = nameResolver.getDisplayName(fullAddress);

          const cellInfo: CellInfo = {
            address: cellAddress,
            fullAddress,
            sheet: sheetName,
            formula,
            value: cell.v,
            references,
            excelName,
          };

          sheetCells.set(cellAddress, cellInfo);
          allCells.set(fullAddress, cellInfo);
        }
      }
    }

    sheets.push({ name: sheetName, cells: sheetCells });
  }

  return {
    fileName: filename,
    sheets,
    allCells,
    definedNames,
  };
}

// =============================================================================
// Test: Global Named Cell
// =============================================================================
describe("Global Named Cell", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("global-named-cell.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "global-named-cell.xlsx");
  });

  it("should extract the defined name 'mass'", () => {
    expect(workbookData.definedNames).toHaveLength(1);
    expect(workbookData.definedNames[0].name).toBe("mass");
    expect(workbookData.definedNames[0].ref).toBe("Sheet1!$A$1");
    expect(workbookData.definedNames[0].sheetScope).toBeUndefined();
  });

  it("should set excelName on the named cell", () => {
    const cell = workbookData.allCells.get("Sheet1!A1");
    expect(cell).toBeDefined();
    expect(cell!.excelName).toBe("mass");
  });

  it("should create edge from formula referencing named cell", () => {
    const cellB1 = workbookData.allCells.get("Sheet1!B1");
    expect(cellB1).toBeDefined();
    expect(cellB1!.formula).toBe("=mass*2");
    expect(cellB1!.references).toContain("Sheet1!A1");
  });

  it("should create edges in graph data for named references", () => {
    const graphData = workbookToGraphData(workbookData);

    // Find edge from A1 to B1 (mass*2)
    const edgeToB1 = graphData.edges.find(
      (e) => e.source === "Sheet1!A1" && e.target === "Sheet1!B1"
    );
    expect(edgeToB1).toBeDefined();

    // Find edge from A1 to A2 (mass+10)
    const edgeToA2 = graphData.edges.find(
      (e) => e.source === "Sheet1!A1" && e.target === "Sheet1!A2"
    );
    expect(edgeToA2).toBeDefined();
  });

  it("should include excelName in graph nodes", () => {
    const graphData = workbookToGraphData(workbookData);
    const nodeA1 = graphData.nodes.find((n) => n.id === "Sheet1!A1");
    expect(nodeA1).toBeDefined();
    expect(nodeA1!.excelName).toBe("mass");
  });
});

// =============================================================================
// Test: Global Named Range
// =============================================================================
describe("Global Named Range", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("global-named-range.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "global-named-range.xlsx");
  });

  it("should extract the defined name 'myRange'", () => {
    expect(workbookData.definedNames).toHaveLength(1);
    expect(workbookData.definedNames[0].name).toBe("myRange");
    expect(workbookData.definedNames[0].ref).toBe("Sheet1!$A$1:$A$3");
  });

  it("should create edges to all cells in the named range", () => {
    const cellA4 = workbookData.allCells.get("Sheet1!A4");
    expect(cellA4).toBeDefined();
    expect(cellA4!.formula).toBe("=SUM(myRange)");

    // Should reference all cells in the range A1:A3
    expect(cellA4!.references).toContain("Sheet1!A1");
    expect(cellA4!.references).toContain("Sheet1!A2");
    expect(cellA4!.references).toContain("Sheet1!A3");
  });

  it("should create graph edges for all range cells", () => {
    const graphData = workbookToGraphData(workbookData);

    // All three cells should have edges to A4
    expect(graphData.edges.some((e) => e.source === "Sheet1!A1" && e.target === "Sheet1!A4")).toBe(true);
    expect(graphData.edges.some((e) => e.source === "Sheet1!A2" && e.target === "Sheet1!A4")).toBe(true);
    expect(graphData.edges.some((e) => e.source === "Sheet1!A3" && e.target === "Sheet1!A4")).toBe(true);
  });

  it("should NOT set excelName on range cells (only single cells get names)", () => {
    // Named ranges don't set excelName on individual cells
    const cellA1 = workbookData.allCells.get("Sheet1!A1");
    expect(cellA1?.excelName).toBeUndefined();
  });
});

// =============================================================================
// Test: Sheet-Scoped Name (Non-Ambiguous)
// =============================================================================
describe("Sheet-Scoped Name (Non-Ambiguous)", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("sheet-scoped-name.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "sheet-scoped-name.xlsx");
  });

  it("should extract sheet-scoped name with correct scope", () => {
    expect(workbookData.definedNames).toHaveLength(1);
    expect(workbookData.definedNames[0].name).toBe("localData");
    expect(workbookData.definedNames[0].sheetScope).toBe(0); // Scoped to Sheet1
  });

  it("should resolve scoped name from same sheet", () => {
    const cellB1 = workbookData.allCells.get("Sheet1!B1");
    expect(cellB1).toBeDefined();
    expect(cellB1!.formula).toBe("=localData+1");
    expect(cellB1!.references).toContain("Sheet1!A1");
  });

  it("should NOT show sheet qualifier for non-ambiguous scoped name", () => {
    const cellA1 = workbookData.allCells.get("Sheet1!A1");
    expect(cellA1).toBeDefined();
    // Since there's only one "localData" name (even though scoped), it shouldn't need qualifier
    expect(cellA1!.excelName).toBe("localData");
  });
});

// =============================================================================
// Test: Ambiguous Names (Same Name in Multiple Scopes)
// =============================================================================
describe("Ambiguous Names", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("ambiguous-names.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "ambiguous-names.xlsx");
  });

  it("should extract both scoped names", () => {
    expect(workbookData.definedNames).toHaveLength(2);
    const names = workbookData.definedNames.map((n) => n.name);
    expect(names).toContain("data");
    expect(names.filter((n) => n === "data")).toHaveLength(2);
  });

  it("should show sheet qualifier for ambiguous names", () => {
    const cellSheet1A1 = workbookData.allCells.get("Sheet1!A1");
    const cellSheet2A1 = workbookData.allCells.get("Sheet2!A1");

    expect(cellSheet1A1).toBeDefined();
    expect(cellSheet2A1).toBeDefined();

    // Both should have qualified names since "data" is ambiguous
    expect(cellSheet1A1!.excelName).toBe("Sheet1::data");
    expect(cellSheet2A1!.excelName).toBe("Sheet2::data");
  });

  it("should resolve to correct cell based on current sheet", () => {
    // Sheet1!B1 formula "data*2" should reference Sheet1!A1
    const cellSheet1B1 = workbookData.allCells.get("Sheet1!B1");
    expect(cellSheet1B1!.references).toContain("Sheet1!A1");
    expect(cellSheet1B1!.references).not.toContain("Sheet2!A1");

    // Sheet2!B1 formula "data*3" should reference Sheet2!A1
    const cellSheet2B1 = workbookData.allCells.get("Sheet2!B1");
    expect(cellSheet2B1!.references).toContain("Sheet2!A1");
    expect(cellSheet2B1!.references).not.toContain("Sheet1!A1");
  });

  it("should create correct edges in graph", () => {
    const graphData = workbookToGraphData(workbookData);

    // Sheet1!A1 -> Sheet1!B1
    expect(graphData.edges.some((e) => e.source === "Sheet1!A1" && e.target === "Sheet1!B1")).toBe(true);

    // Sheet2!A1 -> Sheet2!B1
    expect(graphData.edges.some((e) => e.source === "Sheet2!A1" && e.target === "Sheet2!B1")).toBe(true);

    // Should NOT have cross-sheet edges for these names
    expect(graphData.edges.some((e) => e.source === "Sheet1!A1" && e.target === "Sheet2!B1")).toBe(false);
    expect(graphData.edges.some((e) => e.source === "Sheet2!A1" && e.target === "Sheet1!B1")).toBe(false);
  });
});

// =============================================================================
// Test: Global and Scoped Name with Same Name
// =============================================================================
describe("Global and Scoped Same Name", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("global-and-scoped-same-name.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "global-and-scoped-same-name.xlsx");
  });

  it("should prefer scoped name over global when on same sheet", () => {
    // Sheet1!B1 uses "value" - should resolve to Sheet1!A1 (scoped) not Sheet2!A1 (global)
    const cellSheet1B1 = workbookData.allCells.get("Sheet1!B1");
    expect(cellSheet1B1).toBeDefined();
    expect(cellSheet1B1!.references).toContain("Sheet1!A1");
    expect(cellSheet1B1!.references).not.toContain("Sheet2!A1");
  });

  it("should use global name when no scoped name exists on current sheet", () => {
    // Sheet2!B1 uses "value" - should resolve to Sheet2!A1 (global target)
    const cellSheet2B1 = workbookData.allCells.get("Sheet2!B1");
    expect(cellSheet2B1).toBeDefined();
    expect(cellSheet2B1!.references).toContain("Sheet2!A1");
  });
});

// =============================================================================
// Test: Names with Special Characters
// =============================================================================
describe("Names with Special Characters", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("names-with-special-chars.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "names-with-special-chars.xlsx");
  });

  it("should extract names with underscores and dots", () => {
    const names = workbookData.definedNames.map((n) => n.name);
    expect(names).toContain("my_value");
    expect(names).toContain("tax.rate");
    expect(names).toContain("data_2024");
  });

  it("should resolve all special character names in formula", () => {
    const cellA2 = workbookData.allCells.get("Sheet1!A2");
    expect(cellA2).toBeDefined();
    expect(cellA2!.formula).toBe("=my_value+tax.rate+data_2024");

    // Should have references to all three named cells
    expect(cellA2!.references).toContain("Sheet1!A1"); // my_value
    expect(cellA2!.references).toContain("Sheet1!B1"); // tax.rate
    expect(cellA2!.references).toContain("Sheet1!C1"); // data_2024
  });

  it("should set excelName for cells with special character names", () => {
    expect(workbookData.allCells.get("Sheet1!A1")?.excelName).toBe("my_value");
    expect(workbookData.allCells.get("Sheet1!B1")?.excelName).toBe("tax.rate");
    expect(workbookData.allCells.get("Sheet1!C1")?.excelName).toBe("data_2024");
  });
});

// =============================================================================
// Test: Quoted Sheet Name with Spaces
// =============================================================================
describe("Quoted Sheet Name", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("quoted-sheet-name.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "quoted-sheet-name.xlsx");
  });

  it("should extract name pointing to sheet with spaces", () => {
    expect(workbookData.definedNames).toHaveLength(1);
    expect(workbookData.definedNames[0].name).toBe("revenue");
    // The ref includes quotes around sheet name with spaces
    expect(workbookData.definedNames[0].ref).toContain("My Data Sheet");
  });

  it("should resolve name to correct cell in quoted sheet", () => {
    // Calculations!A1 has formula =revenue*2
    const cellCalcA1 = workbookData.allCells.get("Calculations!A1");
    expect(cellCalcA1).toBeDefined();
    expect(cellCalcA1!.references).toContain("My Data Sheet!A1");
  });

  it("should set excelName on cell in quoted sheet", () => {
    const cellMyDataA1 = workbookData.allCells.get("My Data Sheet!A1");
    expect(cellMyDataA1).toBeDefined();
    expect(cellMyDataA1!.excelName).toBe("revenue");
  });

  it("should create cross-sheet edge in graph", () => {
    const graphData = workbookToGraphData(workbookData);
    const edge = graphData.edges.find(
      (e) => e.source === "My Data Sheet!A1" && e.target === "Calculations!A1"
    );
    expect(edge).toBeDefined();
  });
});

// =============================================================================
// Test: Multiple Names Pointing to Same Cell
// =============================================================================
describe("Multiple Names Same Cell", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("multiple-names-same-cell.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "multiple-names-same-cell.xlsx");
  });

  it("should extract all names for same cell", () => {
    const names = workbookData.definedNames.map((n) => n.name);
    expect(names).toContain("pi");
    expect(names).toContain("circumference");
  });

  it("should prefer one name for display (global preferred)", () => {
    const cellA1 = workbookData.allCells.get("Sheet1!A1");
    expect(cellA1).toBeDefined();
    // Should have one of the names (implementation picks one)
    expect(["pi", "circumference"]).toContain(cellA1!.excelName);
  });

  it("should resolve both names to same cell", () => {
    const cellB1 = workbookData.allCells.get("Sheet1!B1"); // =pi*radius
    const cellC1 = workbookData.allCells.get("Sheet1!C1"); // =circumference

    // Both should reference A1
    expect(cellB1!.references).toContain("Sheet1!A1");
    expect(cellC1!.references).toContain("Sheet1!A1");
  });
});

// =============================================================================
// Test: Cross-Sheet Named Reference
// =============================================================================
describe("Cross-Sheet Named Reference", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("cross-sheet-named-ref.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "cross-sheet-named-ref.xlsx");
  });

  it("should resolve named ref to cell in different sheet", () => {
    // Calculations!A1 has formula =baseValue*1.1
    const cellCalcA1 = workbookData.allCells.get("Calculations!A1");
    expect(cellCalcA1).toBeDefined();
    expect(cellCalcA1!.references).toContain("Data!A1");
  });

  it("should create cross-sheet edge for named reference", () => {
    const graphData = workbookToGraphData(workbookData);
    const namedEdge = graphData.edges.find(
      (e) => e.source === "Data!A1" && e.target === "Calculations!A1"
    );
    expect(namedEdge).toBeDefined();
  });

  it("should also create edge for direct cell reference", () => {
    // Calculations!A2 has formula =Data!A1*1.2 (direct reference)
    const graphData = workbookToGraphData(workbookData);
    const directEdge = graphData.edges.find(
      (e) => e.source === "Data!A1" && e.target === "Calculations!A2"
    );
    expect(directEdge).toBeDefined();
  });
});

// =============================================================================
// Test: Complex Scenario
// =============================================================================
describe("Complex Scenario", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("complex-scenario.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "complex-scenario.xlsx");
  });

  it("should extract all defined names including scoped", () => {
    const names = workbookData.definedNames.map((n) => n.name);
    expect(names).toContain("principal");
    expect(names).toContain("rate");
    expect(names).toContain("months");
    expect(names).toContain("localMultiplier");
  });

  it("should correctly scope localMultiplier to Calculations sheet", () => {
    const localMultiplier = workbookData.definedNames.find((n) => n.name === "localMultiplier");
    expect(localMultiplier).toBeDefined();
    expect(localMultiplier!.sheetScope).toBe(1); // Calculations is index 1
  });

  it("should resolve compound formula with multiple names", () => {
    // Calculations!A1 = principal*(1+rate)^months
    const cellCalcA1 = workbookData.allCells.get("Calculations!A1");
    expect(cellCalcA1).toBeDefined();
    expect(cellCalcA1!.references).toContain("Inputs!A1"); // principal
    expect(cellCalcA1!.references).toContain("Inputs!B1"); // rate
    expect(cellCalcA1!.references).toContain("Inputs!C1"); // months
  });

  it("should resolve mixed named and direct references", () => {
    // Calculations!A2 = principal+Inputs!B1*100
    const cellCalcA2 = workbookData.allCells.get("Calculations!A2");
    expect(cellCalcA2).toBeDefined();
    expect(cellCalcA2!.references).toContain("Inputs!A1"); // principal (named)
    expect(cellCalcA2!.references).toContain("Inputs!B1"); // direct ref
  });

  it("should resolve scoped name from same sheet", () => {
    // Calculations!A3 = localMultiplier*principal
    const cellCalcA3 = workbookData.allCells.get("Calculations!A3");
    expect(cellCalcA3).toBeDefined();
    expect(cellCalcA3!.references).toContain("Calculations!A4"); // localMultiplier
    expect(cellCalcA3!.references).toContain("Inputs!A1"); // principal
  });

  it("should resolve global names from third sheet", () => {
    // Summary!A1 = principal
    const cellSumA1 = workbookData.allCells.get("Summary!A1");
    expect(cellSumA1).toBeDefined();
    expect(cellSumA1!.references).toContain("Inputs!A1");
  });
});

// =============================================================================
// Test: Name Like Cell Reference
// =============================================================================
describe("Name Like Cell Reference", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("name-like-cell-ref.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "name-like-cell-ref.xlsx");
  });

  it("should extract names that look like cell refs with suffix", () => {
    const names = workbookData.definedNames.map((n) => n.name);
    expect(names).toContain("A1_data");
    expect(names).toContain("B1_data");
  });

  it("should resolve names correctly without confusion with cell refs", () => {
    // A2 = A1_data+B1_data
    const cellA2 = workbookData.allCells.get("Sheet1!A2");
    expect(cellA2).toBeDefined();

    // Should reference A1 and B1 via the names
    expect(cellA2!.references).toContain("Sheet1!A1");
    expect(cellA2!.references).toContain("Sheet1!B1");
  });
});

// =============================================================================
// Test: No Defined Names (Control Case)
// =============================================================================
describe("No Defined Names", () => {
  let workbookData: WorkbookData;

  beforeAll(() => {
    const { workbook, definedNames } = loadWorkbookWithNames("no-defined-names.xlsx");
    workbookData = parseWorkbookToData(workbook, definedNames, "no-defined-names.xlsx");
  });

  it("should have empty definedNames array", () => {
    expect(workbookData.definedNames).toHaveLength(0);
  });

  it("should not set excelName on any cells", () => {
    for (const cell of workbookData.allCells.values()) {
      expect(cell.excelName).toBeUndefined();
    }
  });

  it("should still resolve direct cell references", () => {
    const cellB1 = workbookData.allCells.get("Sheet1!B1");
    expect(cellB1).toBeDefined();
    expect(cellB1!.formula).toBe("=A1*2");
    expect(cellB1!.references).toContain("Sheet1!A1");
  });

  it("should create correct edges for direct references", () => {
    const graphData = workbookToGraphData(workbookData);

    // A1 -> B1
    expect(graphData.edges.some((e) => e.source === "Sheet1!A1" && e.target === "Sheet1!B1")).toBe(true);

    // A1, B1 -> A2
    expect(graphData.edges.some((e) => e.source === "Sheet1!A1" && e.target === "Sheet1!A2")).toBe(true);
    expect(graphData.edges.some((e) => e.source === "Sheet1!B1" && e.target === "Sheet1!A2")).toBe(true);
  });
});
