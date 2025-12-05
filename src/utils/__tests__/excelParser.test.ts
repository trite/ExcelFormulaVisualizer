import { describe, it, expect } from "vitest";
import * as XLSX from "xlsx";
import {
  workbookToGraphData,
  graphDataToForceGraph,
  extractMetadataFromWorkbook,
  addMetadataToWorkbook,
} from "../excelParser";
import type {
  WorkbookData,
  SheetData,
  CellInfo,
  CellMetadata,
} from "../../types";

// Helper to create a CellInfo
function createCellInfo(
  address: string,
  sheet: string,
  options: {
    formula?: string;
    value?: unknown;
    references?: string[];
  } = {}
): CellInfo {
  const { formula, value, references = [] } = options;
  return {
    address,
    fullAddress: `${sheet}!${address}`,
    sheet,
    formula,
    value,
    references,
  };
}

// Helper to create a SheetData
function createSheetData(name: string, cells: CellInfo[]): SheetData {
  const cellMap = new Map<string, CellInfo>();
  cells.forEach((cell) => cellMap.set(cell.address, cell));
  return { name, cells: cellMap };
}

// Helper to create a WorkbookData
function createWorkbookData(
  fileName: string,
  sheets: SheetData[]
): WorkbookData {
  const allCells = new Map<string, CellInfo>();
  sheets.forEach((sheet) => {
    sheet.cells.forEach((cell, _address) => {
      allCells.set(cell.fullAddress, cell);
    });
  });
  return { fileName, sheets, allCells };
}

describe("excelParser", () => {
  describe("workbookToGraphData - Edge Generation", () => {
    it("should create edges from formula references", () => {
      // A1 has value, B1 references A1
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 100 }),
        createCellInfo("B1", "Sheet1", {
          formula: "=A1*2",
          value: 200,
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);

      expect(graphData.edges).toHaveLength(1);
      expect(graphData.edges[0]).toEqual({
        id: "Sheet1!A1->Sheet1!B1",
        source: "Sheet1!A1",
        target: "Sheet1!B1",
      });
    });

    it("should create edge with source as referenced cell and target as formula cell", () => {
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 10 }),
        createCellInfo("A2", "Sheet1", {
          formula: "=A1+5",
          value: 15,
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);

      const edge = graphData.edges[0];
      expect(edge.source).toBe("Sheet1!A1"); // Referenced cell
      expect(edge.target).toBe("Sheet1!A2"); // Formula cell
    });

    it("should create cross-sheet edges correctly", () => {
      const sheet1Cells = [createCellInfo("A1", "Sheet1", { value: 100 })];
      const sheet2Cells = [
        createCellInfo("A1", "Sheet2", {
          formula: "=Sheet1!A1*2",
          value: 200,
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet1 = createSheetData("Sheet1", sheet1Cells);
      const sheet2 = createSheetData("Sheet2", sheet2Cells);
      const workbook = createWorkbookData("test.xlsx", [sheet1, sheet2]);
      const graphData = workbookToGraphData(workbook);

      expect(graphData.edges).toHaveLength(1);
      expect(graphData.edges[0]).toEqual({
        id: "Sheet1!A1->Sheet2!A1",
        source: "Sheet1!A1",
        target: "Sheet2!A1",
      });
    });

    it("should create multiple edges for formula with multiple references", () => {
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 10 }),
        createCellInfo("A2", "Sheet1", { value: 20 }),
        createCellInfo("A3", "Sheet1", {
          formula: "=A1+A2",
          value: 30,
          references: ["Sheet1!A1", "Sheet1!A2"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);

      expect(graphData.edges).toHaveLength(2);

      const edgeSources = graphData.edges.map((e) => e.source).sort();
      expect(edgeSources).toEqual(["Sheet1!A1", "Sheet1!A2"]);

      // Both edges should target A3
      expect(graphData.edges.every((e) => e.target === "Sheet1!A3")).toBe(true);
    });

    it("should create chain of edges for dependent formulas", () => {
      // A1 -> A2 -> A3 (chain of dependencies)
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 10 }),
        createCellInfo("A2", "Sheet1", {
          formula: "=A1*2",
          value: 20,
          references: ["Sheet1!A1"],
        }),
        createCellInfo("A3", "Sheet1", {
          formula: "=A2+5",
          value: 25,
          references: ["Sheet1!A2"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);

      expect(graphData.edges).toHaveLength(2);

      // Edge from A1 to A2
      expect(graphData.edges).toContainEqual({
        id: "Sheet1!A1->Sheet1!A2",
        source: "Sheet1!A1",
        target: "Sheet1!A2",
      });

      // Edge from A2 to A3
      expect(graphData.edges).toContainEqual({
        id: "Sheet1!A2->Sheet1!A3",
        source: "Sheet1!A2",
        target: "Sheet1!A3",
      });
    });

    it("should match expected edge structure from edge-test fixture", () => {
      // This simulates the edge-test.xlsx structure
      const sheet1Cells = [
        createCellInfo("A1", "Sheet1", { value: 10 }),
        createCellInfo("A2", "Sheet1", { value: 20 }),
        createCellInfo("A3", "Sheet1", {
          formula: "=A1+A2",
          value: 30,
          references: ["Sheet1!A1", "Sheet1!A2"],
        }),
        createCellInfo("B1", "Sheet1", {
          formula: "=A3*2",
          value: 60,
          references: ["Sheet1!A3"],
        }),
      ];

      const sheet2Cells = [
        createCellInfo("A1", "Sheet2", {
          formula: "=Sheet1!A1+Sheet1!A2",
          value: 30,
          references: ["Sheet1!A1", "Sheet1!A2"],
        }),
        createCellInfo("A2", "Sheet2", {
          formula: "=Sheet1!B1",
          value: 60,
          references: ["Sheet1!B1"],
        }),
      ];

      const sheet1 = createSheetData("Sheet1", sheet1Cells);
      const sheet2 = createSheetData("Sheet2", sheet2Cells);
      const workbook = createWorkbookData("test.xlsx", [sheet1, sheet2]);
      const graphData = workbookToGraphData(workbook);

      // Expected edges:
      // 1. Sheet1!A1 -> Sheet1!A3
      // 2. Sheet1!A2 -> Sheet1!A3
      // 3. Sheet1!A3 -> Sheet1!B1
      // 4. Sheet1!A1 -> Sheet2!A1
      // 5. Sheet1!A2 -> Sheet2!A1
      // 6. Sheet1!B1 -> Sheet2!A2
      expect(graphData.edges).toHaveLength(6);

      const expectedEdges = [
        { source: "Sheet1!A1", target: "Sheet1!A3" },
        { source: "Sheet1!A2", target: "Sheet1!A3" },
        { source: "Sheet1!A3", target: "Sheet1!B1" },
        { source: "Sheet1!A1", target: "Sheet2!A1" },
        { source: "Sheet1!A2", target: "Sheet2!A1" },
        { source: "Sheet1!B1", target: "Sheet2!A2" },
      ];

      expectedEdges.forEach((expected) => {
        const found = graphData.edges.find(
          (e) => e.source === expected.source && e.target === expected.target
        );
        expect(
          found,
          `Expected edge ${expected.source} -> ${expected.target}`
        ).toBeDefined();
      });
    });
  });

  describe("workbookToGraphData - Node Generation", () => {
    it("should create nodes for cells with formulas", () => {
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 100 }),
        createCellInfo("B1", "Sheet1", {
          formula: "=A1*2",
          value: 200,
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);

      const formulaNode = graphData.nodes.find((n) => n.id === "Sheet1!B1");
      expect(formulaNode).toBeDefined();
      expect(formulaNode?.hasFormula).toBe(true);
    });

    it("should create nodes for referenced cells even without formulas", () => {
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 100 }),
        createCellInfo("B1", "Sheet1", {
          formula: "=A1*2",
          value: 200,
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);

      const referencedNode = graphData.nodes.find((n) => n.id === "Sheet1!A1");
      expect(referencedNode).toBeDefined();
      expect(referencedNode?.hasFormula).toBe(false);
    });

    it("should set correct hasFormula flag", () => {
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 100 }),
        createCellInfo("B1", "Sheet1", {
          formula: "=A1*2",
          value: 200,
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);

      const valueNode = graphData.nodes.find((n) => n.id === "Sheet1!A1");
      const formulaNode = graphData.nodes.find((n) => n.id === "Sheet1!B1");

      expect(valueNode?.hasFormula).toBe(false);
      expect(formulaNode?.hasFormula).toBe(true);
    });

    it("should use full address as node id", () => {
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 100 }),
        createCellInfo("B1", "Sheet1", {
          formula: "=A1",
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);

      expect(graphData.nodes.every((n) => n.id.includes("!"))).toBe(true);
    });

    it("should not create duplicate nodes", () => {
      // A1 is referenced by both B1 and C1
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 10 }),
        createCellInfo("B1", "Sheet1", {
          formula: "=A1*2",
          references: ["Sheet1!A1"],
        }),
        createCellInfo("C1", "Sheet1", {
          formula: "=A1+5",
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);

      const a1Nodes = graphData.nodes.filter((n) => n.id === "Sheet1!A1");
      expect(a1Nodes).toHaveLength(1);
    });
  });

  describe("graphDataToForceGraph", () => {
    it("should convert graph data to force graph format", () => {
      const cells = [
        createCellInfo("A1", "Sheet1", { value: 100 }),
        createCellInfo("B1", "Sheet1", {
          formula: "=A1*2",
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet = createSheetData("Sheet1", cells);
      const workbook = createWorkbookData("test.xlsx", [sheet]);
      const graphData = workbookToGraphData(workbook);
      const forceGraph = graphDataToForceGraph(graphData, workbook.sheets);

      expect(forceGraph.nodes).toHaveLength(graphData.nodes.length);
      expect(forceGraph.links).toHaveLength(graphData.edges.length);

      // Links should have source and target
      expect(forceGraph.links[0]).toHaveProperty("source");
      expect(forceGraph.links[0]).toHaveProperty("target");
    });

    it("should assign colors to nodes based on sheet", () => {
      const sheet1Cells = [
        createCellInfo("A1", "Sheet1", { value: 100 }),
        createCellInfo("B1", "Sheet1", {
          formula: "=A1",
          references: ["Sheet1!A1"],
        }),
      ];
      const sheet2Cells = [
        createCellInfo("A1", "Sheet2", {
          formula: "=Sheet1!A1",
          references: ["Sheet1!A1"],
        }),
      ];

      const sheet1 = createSheetData("Sheet1", sheet1Cells);
      const sheet2 = createSheetData("Sheet2", sheet2Cells);
      const workbook = createWorkbookData("test.xlsx", [sheet1, sheet2]);
      const graphData = workbookToGraphData(workbook);
      const forceGraph = graphDataToForceGraph(graphData, workbook.sheets);

      // All nodes should have colors
      expect(forceGraph.nodes.every((n) => n.color !== undefined)).toBe(true);

      // Nodes from same sheet should have same color
      const sheet1Nodes = forceGraph.nodes.filter((n) => n.sheet === "Sheet1");
      const sheet1Color = sheet1Nodes[0]?.color;
      expect(sheet1Nodes.every((n) => n.color === sheet1Color)).toBe(true);

      // Nodes from different sheets should have different colors
      const sheet2Node = forceGraph.nodes.find((n) => n.sheet === "Sheet2");
      expect(sheet2Node?.color).not.toBe(sheet1Color);
    });
  });

  describe("extractMetadataFromWorkbook", () => {
    it("should return null when no metadata sheet exists", () => {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([
        [1, 2],
        [3, 4],
      ]);
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      const metadata = extractMetadataFromWorkbook(wb);
      expect(metadata).toBeNull();
    });

    it("should parse names from columns A/B", () => {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([[1]]);
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      const metadataWs = XLSX.utils.aoa_to_sheet([
        ["Name Keys", "Name Values", "Note Keys", "Note Values"],
        ["Sheet1!A1", "MyName", "", ""],
        ["Sheet1!B1", "OtherName", "", ""],
      ]);
      XLSX.utils.book_append_sheet(
        wb,
        metadataWs,
        "ExcelFormulaVisualizerMetadata"
      );

      const metadata = extractMetadataFromWorkbook(wb);

      expect(metadata).not.toBeNull();
      expect(metadata?.names).toHaveLength(2);
      expect(metadata?.names[0]).toEqual({
        cellKey: "Sheet1!A1",
        name: "MyName",
      });
      expect(metadata?.names[1]).toEqual({
        cellKey: "Sheet1!B1",
        name: "OtherName",
      });
    });

    it("should parse notes from columns C/D", () => {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([[1]]);
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      const metadataWs = XLSX.utils.aoa_to_sheet([
        ["Name Keys", "Name Values", "Note Keys", "Note Values"],
        ["", "", "Sheet1!A1", "First note"],
        ["", "", "Sheet1!A1", "Second note"],
      ]);
      XLSX.utils.book_append_sheet(
        wb,
        metadataWs,
        "ExcelFormulaVisualizerMetadata"
      );

      const metadata = extractMetadataFromWorkbook(wb);

      expect(metadata?.notes).toHaveLength(2);
      expect(metadata?.notes[0]).toEqual({
        cellKey: "Sheet1!A1",
        note: "First note",
      });
      expect(metadata?.notes[1]).toEqual({
        cellKey: "Sheet1!A1",
        note: "Second note",
      });
    });

    it("should skip header row", () => {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([[1]]);
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      const metadataWs = XLSX.utils.aoa_to_sheet([
        ["Name Keys", "Name Values", "Note Keys", "Note Values"],
        ["Sheet1!A1", "ActualName", "", ""],
      ]);
      XLSX.utils.book_append_sheet(
        wb,
        metadataWs,
        "ExcelFormulaVisualizerMetadata"
      );

      const metadata = extractMetadataFromWorkbook(wb);

      // Should not include header as a name
      expect(metadata?.names).toHaveLength(1);
      expect(metadata?.names[0].cellKey).toBe("Sheet1!A1");
      expect(metadata?.names[0].name).toBe("ActualName");
    });
  });

  describe("addMetadataToWorkbook", () => {
    it("should add metadata sheet to workbook", () => {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([[1]]);
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      const metadata: CellMetadata = {
        names: [{ cellKey: "Sheet1!A1", name: "TestName" }],
        notes: [{ cellKey: "Sheet1!A1", note: "TestNote" }],
      };

      addMetadataToWorkbook(wb, metadata);

      expect(wb.SheetNames).toContain("ExcelFormulaVisualizerMetadata");
    });

    it("should set metadata sheet as hidden", () => {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([[1]]);
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      const metadata: CellMetadata = { names: [], notes: [] };
      addMetadataToWorkbook(wb, metadata);

      const metadataIdx = wb.SheetNames.indexOf(
        "ExcelFormulaVisualizerMetadata"
      );
      expect(wb.Workbook?.Sheets?.[metadataIdx]?.Hidden).toBe(1);
    });

    it("should replace existing metadata sheet", () => {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([[1]]);
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      // Add first metadata
      const metadata1: CellMetadata = {
        names: [{ cellKey: "Sheet1!A1", name: "OldName" }],
        notes: [],
      };
      addMetadataToWorkbook(wb, metadata1);

      // Add second metadata (should replace)
      const metadata2: CellMetadata = {
        names: [{ cellKey: "Sheet1!A1", name: "NewName" }],
        notes: [],
      };
      addMetadataToWorkbook(wb, metadata2);

      // Should only have one metadata sheet
      const metadataSheetCount = wb.SheetNames.filter(
        (name) => name === "ExcelFormulaVisualizerMetadata"
      ).length;
      expect(metadataSheetCount).toBe(1);

      // Should have the new metadata
      const extracted = extractMetadataFromWorkbook(wb);
      expect(extracted?.names[0].name).toBe("NewName");
    });
  });
});
