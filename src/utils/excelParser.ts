import * as XLSX from "xlsx";
import type {
  WorkbookData,
  SheetData,
  CellInfo,
  GraphData,
  GraphNode,
  GraphEdge,
  ForceGraphData,
  CellMetadata,
  DefinedNameInfo,
} from "../types";
import { extractCellReferences } from "./formulaParser";
import { createNameResolver } from "./nameResolver";

const METADATA_SHEET_NAME = "ExcelFormulaVisualizerMetadata";

/**
 * Extract defined names from an Excel workbook
 * These are names created via Excel's Name Manager (Ctrl+F3)
 */
function extractDefinedNames(workbook: XLSX.WorkBook): DefinedNameInfo[] {
  const definedNames: DefinedNameInfo[] = [];

  // xlsx stores defined names in workbook.Workbook.Names
  const names = workbook.Workbook?.Names;
  if (!names || !Array.isArray(names)) {
    return definedNames;
  }

  for (const name of names) {
    // Skip built-in names (like _xlnm.Print_Area)
    if (!name.Name || name.Name.startsWith("_xlnm.")) {
      continue;
    }

    // Skip names with invalid references
    if (!name.Ref || name.Ref.startsWith("#")) {
      continue;
    }

    definedNames.push({
      name: name.Name,
      ref: name.Ref,
      sheetScope: name.Sheet, // undefined for global, 0-indexed for sheet-scoped
      comment: name.Comment,
    });
  }

  return definedNames;
}

/**
 * Parses an Excel file and extracts all cell data with formulas
 */
export async function parseExcelFile(file: File): Promise<WorkbookData> {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { cellFormula: true, cellNF: true });

  // Extract defined names and create resolver
  const definedNames = extractDefinedNames(workbook);
  const sheetNames = workbook.SheetNames;
  const nameResolver = createNameResolver(definedNames, sheetNames);

  const sheets: SheetData[] = [];
  const allCells = new Map<string, CellInfo>();

  for (let sheetIndex = 0; sheetIndex < sheetNames.length; sheetIndex++) {
    const sheetName = sheetNames[sheetIndex];
    const worksheet = workbook.Sheets[sheetName];
    const sheetCells = new Map<string, CellInfo>();

    // Get the range of the sheet
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

          // Check if this cell has an Excel-defined name
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

    sheets.push({
      name: sheetName,
      cells: sheetCells,
    });
  }

  return {
    fileName: file.name,
    sheets,
    allCells,
    definedNames,
  };
}

/**
 * Generates a color for a sheet based on its index
 */
function getSheetColor(sheetIndex: number, totalSheets: number): string {
  const hue = (sheetIndex / totalSheets) * 360;
  return `hsl(${hue}, 70%, 50%)`;
}

/**
 * Converts workbook data to graph data structure for visualization
 */
export function workbookToGraphData(workbook: WorkbookData): GraphData {
  const nodes: GraphNode[] = [];
  const edges: GraphEdge[] = [];
  const nodeIds = new Set<string>();
  const sheetIndices = new Map<string, number>();

  // Create sheet index mapping for coloring
  workbook.sheets.forEach((sheet, index) => {
    sheetIndices.set(sheet.name, index);
  });

  // First pass: create nodes for all cells that have formulas or are referenced
  workbook.allCells.forEach((cell, fullAddress) => {
    if (cell.formula || cell.references.length > 0) {
      if (!nodeIds.has(fullAddress)) {
        nodes.push({
          id: fullAddress,
          label: cell.address,
          sheet: cell.sheet,
          address: cell.address,
          formula: cell.formula,
          value: cell.value,
          hasFormula: !!cell.formula,
          excelName: cell.excelName,
        });
        nodeIds.add(fullAddress);
      }

      // Create nodes for referenced cells
      cell.references.forEach((refAddress) => {
        if (!nodeIds.has(refAddress)) {
          const refCell = workbook.allCells.get(refAddress);
          const [refSheet, refCellAddr] = refAddress.split("!");
          nodes.push({
            id: refAddress,
            label: refCellAddr || refAddress,
            sheet: refSheet || cell.sheet,
            address: refCellAddr || refAddress,
            formula: refCell?.formula,
            value: refCell?.value,
            hasFormula: !!refCell?.formula,
            excelName: refCell?.excelName,
          });
          nodeIds.add(refAddress);
        }
      });
    }
  });

  // Second pass: create edges
  workbook.allCells.forEach((cell) => {
    if (cell.formula && cell.references.length > 0) {
      cell.references.forEach((refAddress) => {
        edges.push({
          id: `${refAddress}->${cell.fullAddress}`,
          source: refAddress,
          target: cell.fullAddress,
        });
      });
    }
  });

  return { nodes, edges };
}

/**
 * Converts graph data to Force-Graph format
 */
export function graphDataToForceGraph(
  graphData: GraphData,
  sheets: SheetData[]
): ForceGraphData {
  const sheetIndices = new Map<string, number>();
  sheets.forEach((sheet, index) => {
    sheetIndices.set(sheet.name, index);
  });

  const nodes = graphData.nodes.map((node) => ({
    ...node,
    color: getSheetColor(sheetIndices.get(node.sheet) || 0, sheets.length || 1),
  }));

  const links = graphData.edges.map((edge) => ({
    source: edge.source,
    target: edge.target,
  }));

  return { nodes, links };
}

/**
 * Extract metadata from the ExcelFormulaVisualizerMetadata sheet if it exists
 */
export function extractMetadataFromWorkbook(
  workbook: XLSX.WorkBook
): CellMetadata | null {
  if (!workbook.SheetNames.includes(METADATA_SHEET_NAME)) {
    return null;
  }

  const worksheet = workbook.Sheets[METADATA_SHEET_NAME];
  if (!worksheet || !worksheet["!ref"]) {
    return null;
  }

  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  const metadata: CellMetadata = { names: [], notes: [] };

  // Skip header row (row 0), start from row 1
  for (let row = 1; row <= range.e.r; row++) {
    // Column A: Name keys, Column B: Name values
    const nameKeyCell = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
    const nameValueCell = worksheet[XLSX.utils.encode_cell({ r: row, c: 1 })];

    if (nameKeyCell?.v && nameValueCell?.v) {
      metadata.names.push({
        cellKey: String(nameKeyCell.v),
        name: String(nameValueCell.v),
      });
    }

    // Column C: Note keys, Column D: Note values
    const noteKeyCell = worksheet[XLSX.utils.encode_cell({ r: row, c: 2 })];
    const noteValueCell = worksheet[XLSX.utils.encode_cell({ r: row, c: 3 })];

    if (noteKeyCell?.v && noteValueCell?.v) {
      metadata.notes.push({
        cellKey: String(noteKeyCell.v),
        note: String(noteValueCell.v),
      });
    }
  }

  return metadata;
}

/**
 * Parse an Excel file and also extract any existing metadata
 */
export async function parseExcelFileWithMetadata(file: File): Promise<{
  workbookData: WorkbookData;
  excelMetadata: CellMetadata | null;
  rawWorkbook: XLSX.WorkBook;
}> {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { cellFormula: true, cellNF: true });

  // Extract metadata before processing sheets
  const excelMetadata = extractMetadataFromWorkbook(workbook);

  // Extract defined names and create resolver
  const definedNames = extractDefinedNames(workbook);
  // Filter out metadata sheet from sheet names for the resolver
  const sheetNames = workbook.SheetNames.filter(
    (name) => name !== METADATA_SHEET_NAME
  );
  const nameResolver = createNameResolver(definedNames, sheetNames);

  const sheets: SheetData[] = [];
  const allCells = new Map<string, CellInfo>();

  // Track sheet index excluding metadata sheet
  let sheetIndex = 0;
  for (const sheetName of workbook.SheetNames) {
    // Skip the metadata sheet
    if (sheetName === METADATA_SHEET_NAME) continue;

    const worksheet = workbook.Sheets[sheetName];
    const sheetCells = new Map<string, CellInfo>();

    // Get the range of the sheet
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

          // Check if this cell has an Excel-defined name
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

    sheets.push({
      name: sheetName,
      cells: sheetCells,
    });

    sheetIndex++;
  }

  return {
    workbookData: {
      fileName: file.name,
      sheets,
      allCells,
      definedNames,
    },
    excelMetadata,
    rawWorkbook: workbook,
  };
}

/**
 * Add metadata to a workbook as a hidden sheet
 */
export function addMetadataToWorkbook(
  workbook: XLSX.WorkBook,
  metadata: CellMetadata
): XLSX.WorkBook {
  // Create the metadata sheet data
  const sheetData: (string | undefined)[][] = [
    ["Name Keys", "Name Values", "Note Keys", "Note Values"], // Header row
  ];

  // Determine max rows needed
  const maxRows = Math.max(metadata.names.length, metadata.notes.length);

  for (let i = 0; i < maxRows; i++) {
    const row: (string | undefined)[] = [
      metadata.names[i]?.cellKey,
      metadata.names[i]?.name,
      metadata.notes[i]?.cellKey,
      metadata.notes[i]?.note,
    ];
    sheetData.push(row);
  }

  // Create worksheet from array
  const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

  // Set column widths for readability
  worksheet["!cols"] = [
    { wch: 30 }, // Name Keys
    { wch: 20 }, // Name Values
    { wch: 30 }, // Note Keys
    { wch: 40 }, // Note Values
  ];

  // Remove existing metadata sheet if present
  const sheetIndex = workbook.SheetNames.indexOf(METADATA_SHEET_NAME);
  if (sheetIndex > -1) {
    workbook.SheetNames.splice(sheetIndex, 1);
    delete workbook.Sheets[METADATA_SHEET_NAME];
  }

  // Add the new metadata sheet
  workbook.SheetNames.push(METADATA_SHEET_NAME);
  workbook.Sheets[METADATA_SHEET_NAME] = worksheet;

  // Hide the metadata sheet
  if (!workbook.Workbook) {
    workbook.Workbook = { Sheets: [] };
  }
  if (!workbook.Workbook.Sheets) {
    workbook.Workbook.Sheets = [];
  }

  // Ensure we have sheet properties for all sheets
  while (workbook.Workbook.Sheets.length < workbook.SheetNames.length) {
    workbook.Workbook.Sheets.push({});
  }

  // Set the metadata sheet as hidden (1 = hidden, 2 = very hidden)
  const metadataSheetIdx = workbook.SheetNames.indexOf(METADATA_SHEET_NAME);
  workbook.Workbook.Sheets[metadataSheetIdx] = { Hidden: 1 };

  return workbook;
}

/**
 * Download a workbook with metadata as an Excel file
 */
export function downloadWorkbookWithMetadata(
  rawWorkbook: XLSX.WorkBook,
  metadata: CellMetadata,
  fileName: string
): void {
  // Clone the workbook to avoid modifying the original
  const workbookCopy = JSON.parse(JSON.stringify(rawWorkbook));

  // Add metadata
  addMetadataToWorkbook(workbookCopy, metadata);

  // Generate filename with metadata suffix if not already present
  let outputName = fileName;
  if (!outputName.includes("_with_metadata")) {
    const lastDot = outputName.lastIndexOf(".");
    if (lastDot > 0) {
      outputName =
        outputName.slice(0, lastDot) +
        "_with_metadata" +
        outputName.slice(lastDot);
    } else {
      outputName = outputName + "_with_metadata.xlsx";
    }
  }

  // Write and download
  XLSX.writeFile(workbookCopy, outputName);
}
