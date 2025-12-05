// Core types for the Excel Formula Visualizer

export interface CellInfo {
  address: string; // e.g., "A1"
  fullAddress: string; // e.g., "Sheet1!A1"
  sheet: string;
  formula?: string;
  value?: unknown;
  references: string[]; // Full addresses of cells this cell references
  excelName?: string; // Excel-defined name for this cell (from Name Manager)
}

export interface SheetData {
  name: string;
  cells: Map<string, CellInfo>;
}

export interface WorkbookData {
  fileName: string;
  sheets: SheetData[];
  allCells: Map<string, CellInfo>; // Keyed by fullAddress
  definedNames: DefinedNameInfo[]; // Excel-defined names from Name Manager
}

// Excel defined name (from Name Manager)
export interface DefinedNameInfo {
  name: string; // The name (e.g., "mass", "myData")
  ref: string; // The reference (e.g., "Sheet1!$A$1" or "Sheet1!$A$1:$B$10")
  sheetScope?: number; // If defined, this name is sheet-scoped (0-indexed sheet)
  comment?: string; // Optional comment from Excel
}

export interface GraphNode {
  id: string; // fullAddress
  label: string;
  sheet: string;
  address: string;
  formula?: string;
  value?: unknown;
  hasFormula: boolean;
  excelName?: string; // Excel-defined name for this cell
}

export interface GraphEdge {
  id: string;
  source: string; // fullAddress of referenced cell
  target: string; // fullAddress of cell containing the formula
}

export interface GraphData {
  nodes: GraphNode[];
  edges: GraphEdge[];
}

// For Force-Graph library
export interface ForceGraphData {
  nodes: ForceGraphNode[];
  links: ForceGraphLink[];
}

export interface ForceGraphNode {
  id: string;
  label: string;
  sheet: string;
  address: string;
  formula?: string;
  value?: unknown;
  hasFormula: boolean;
  excelName?: string; // Excel-defined name for this cell
  color?: string;
  // Runtime properties added by force-graph
  x?: number;
  y?: number;
  fx?: number;
  fy?: number;
}

export interface ForceGraphLink {
  source: string;
  target: string;
}

// Cell metadata for user-defined names and notes
export interface CellNameEntry {
  cellKey: string; // Comma-separated cell list or range (e.g., "Sheet1!A1,Sheet1!B2" or "Sheet1!A1:B5")
  name: string;
}

export interface CellNoteEntry {
  cellKey: string; // Comma-separated cell list or range
  note: string;
}

export interface CellMetadata {
  names: CellNameEntry[];
  notes: CellNoteEntry[];
}

// Serializable version for localStorage
export interface SerializedCellMetadata {
  names: { cellKey: string; name: string }[];
  notes: { cellKey: string; note: string }[];
}

// For conflict resolution when loading
export interface MetadataConflict {
  hasConflict: boolean;
  localStorage: CellMetadata | null;
  excelSheet: CellMetadata | null;
  differences: MetadataDifference[];
}

export interface MetadataDifference {
  type: "name" | "note";
  cellKey: string;
  localValue?: string;
  excelValue?: string;
  status: "only-local" | "only-excel" | "different";
}

// Graph filtering types
export interface GraphFilterOptions {
  maxNodes: number;
  showOnlyFormulas: boolean;
  selectedSheets: string[] | "all";
  prioritizeNamed: boolean;
  namedCells: Set<string>;
}

export interface GraphFilterStats {
  totalNodes: number;
  totalEdges: number;
  visibleNodes: number;
  visibleEdges: number;
  hiddenBySheet: number;
  hiddenByFormula: number;
  hiddenByLimit: number;
  limitReached: boolean;
}

export interface FilteredGraphResult {
  graphData: GraphData;
  stats: GraphFilterStats;
}
