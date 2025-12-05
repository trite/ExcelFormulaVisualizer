// Utility for managing cell metadata persistence in localStorage

import type {
  CellMetadata,
  SerializedCellMetadata,
  CellNameEntry,
  CellNoteEntry,
  MetadataConflict,
  MetadataDifference,
} from "../types";

const STORAGE_PREFIX = "efv_metadata_";

/**
 * Get the localStorage key for a given filename
 */
function getStorageKey(fileName: string): string {
  return `${STORAGE_PREFIX}${fileName}`;
}

/**
 * Create an empty CellMetadata object
 */
export function createEmptyMetadata(): CellMetadata {
  return {
    names: [],
    notes: [],
  };
}

/**
 * Load metadata from localStorage for a given filename
 */
export function loadMetadataFromStorage(fileName: string): CellMetadata | null {
  try {
    const key = getStorageKey(fileName);
    const stored = localStorage.getItem(key);
    if (!stored) return null;

    const parsed: SerializedCellMetadata = JSON.parse(stored);
    return {
      names: parsed.names || [],
      notes: parsed.notes || [],
    };
  } catch (error) {
    console.error("Failed to load metadata from localStorage:", error);
    return null;
  }
}

/**
 * Save metadata to localStorage for a given filename
 */
export function saveMetadataToStorage(
  fileName: string,
  metadata: CellMetadata
): void {
  try {
    const key = getStorageKey(fileName);
    const serialized: SerializedCellMetadata = {
      names: metadata.names,
      notes: metadata.notes,
    };
    localStorage.setItem(key, JSON.stringify(serialized));
  } catch (error) {
    console.error("Failed to save metadata to localStorage:", error);
  }
}

/**
 * Clear metadata from localStorage for a given filename
 */
export function clearMetadataFromStorage(fileName: string): void {
  try {
    const key = getStorageKey(fileName);
    localStorage.removeItem(key);
  } catch (error) {
    console.error("Failed to clear metadata from localStorage:", error);
  }
}

/**
 * Parse a cell reference like "A1" into column and row numbers
 */
function parseCellRef(cellRef: string): { col: number; row: number } | null {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;

  const colStr = match[1].toUpperCase();
  const row = parseInt(match[2], 10);

  // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }

  return { col, row };
}

/**
 * Convert column number back to letters (1=A, 2=B, ..., 27=AA, etc.)
 */
function colToLetters(col: number): string {
  let result = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    col = Math.floor((col - 1) / 26);
  }
  return result;
}

/**
 * Expand a cell range like "A1:B3" into individual cells
 */
function expandRange(
  sheet: string,
  startCell: string,
  endCell: string
): string[] {
  const start = parseCellRef(startCell);
  const end = parseCellRef(endCell);
  if (!start || !end) return [];

  const cells: string[] = [];
  const minCol = Math.min(start.col, end.col);
  const maxCol = Math.max(start.col, end.col);
  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);

  // Limit expansion to prevent memory issues
  const totalCells = (maxCol - minCol + 1) * (maxRow - minRow + 1);
  if (totalCells > 10000) {
    console.warn(`Range too large (${totalCells} cells), skipping expansion`);
    return [];
  }

  for (let row = minRow; row <= maxRow; row++) {
    for (let col = minCol; col <= maxCol; col++) {
      cells.push(`${sheet}!${colToLetters(col)}${row}`);
    }
  }

  return cells;
}

/**
 * Expand a cell key (which may contain multiple cells and ranges) into individual cell IDs
 * Examples:
 *   "Sheet1!A1" -> ["Sheet1!A1"]
 *   "Sheet1!A1,Sheet1!B2" -> ["Sheet1!A1", "Sheet1!B2"]
 *   "Sheet1!A1:B3" -> ["Sheet1!A1", "Sheet1!A2", "Sheet1!A3", "Sheet1!B1", "Sheet1!B2", "Sheet1!B3"]
 *   "'My Sheet'!A1" -> ["My Sheet!A1"]
 */
export function expandCellKey(cellKey: string): string[] {
  const cells: string[] = [];
  const parts = cellKey.split(",").map((p) => p.trim());

  for (const part of parts) {
    // Match: optional quoted sheet name, !, cell reference, optional :endCell
    const rangeMatch = part.match(
      /^'?([^'!]+)'?!([A-Z]+\d+)(?::([A-Z]+\d+))?$/i
    );
    if (rangeMatch) {
      const sheet = rangeMatch[1];
      const startCell = rangeMatch[2].toUpperCase();
      const endCell = rangeMatch[3]?.toUpperCase();

      if (endCell) {
        // It's a range
        cells.push(...expandRange(sheet, startCell, endCell));
      } else {
        // Single cell
        cells.push(`${sheet}!${startCell}`);
      }
    }
  }

  return cells;
}

/**
 * Check if a cell already has a name assigned
 * Returns the existing entry if found
 */
export function findExistingName(
  cellId: string,
  metadata: CellMetadata
): CellNameEntry | null {
  for (const entry of metadata.names) {
    const expandedCells = expandCellKey(entry.cellKey);
    if (expandedCells.includes(cellId)) {
      return entry;
    }
  }
  return null;
}

/**
 * Find all notes that apply to a given cell
 */
export function findNotesForCell(
  cellId: string,
  metadata: CellMetadata
): CellNoteEntry[] {
  return metadata.notes.filter((entry) => {
    const expandedCells = expandCellKey(entry.cellKey);
    return expandedCells.includes(cellId);
  });
}

/**
 * Get the name for a specific cell (if any)
 */
export function getNameForCell(
  cellId: string,
  metadata: CellMetadata
): string | null {
  const entry = findExistingName(cellId, metadata);
  return entry ? entry.name : null;
}

/**
 * Validate that a new name assignment won't conflict with existing names
 * Returns list of conflicting cells (cells that already have names)
 */
export function validateNameAssignment(
  cellKey: string,
  metadata: CellMetadata,
  excludeKey?: string // Exclude this key from validation (for editing existing name)
): string[] {
  const newCells = expandCellKey(cellKey);
  const conflicts: string[] = [];

  for (const entry of metadata.names) {
    // Skip the entry we're editing
    if (excludeKey && entry.cellKey === excludeKey) continue;

    const existingCells = expandCellKey(entry.cellKey);
    for (const cell of newCells) {
      if (existingCells.includes(cell)) {
        conflicts.push(`${cell} already has name "${entry.name}"`);
      }
    }
  }

  return conflicts;
}

/**
 * Add or update a name entry
 */
export function setName(
  metadata: CellMetadata,
  cellKey: string,
  name: string
): CellMetadata {
  // Remove any existing entry with the same key
  const filteredNames = metadata.names.filter((e) => e.cellKey !== cellKey);

  return {
    ...metadata,
    names: [...filteredNames, { cellKey, name }],
  };
}

/**
 * Remove a name entry by its cell key
 */
export function removeName(
  metadata: CellMetadata,
  cellKey: string
): CellMetadata {
  return {
    ...metadata,
    names: metadata.names.filter((e) => e.cellKey !== cellKey),
  };
}

/**
 * Add a new note entry
 */
export function addNote(
  metadata: CellMetadata,
  cellKey: string,
  note: string
): CellMetadata {
  return {
    ...metadata,
    notes: [...metadata.notes, { cellKey, note }],
  };
}

/**
 * Update an existing note entry
 */
export function updateNote(
  metadata: CellMetadata,
  oldCellKey: string,
  oldNote: string,
  newNote: string
): CellMetadata {
  return {
    ...metadata,
    notes: metadata.notes.map((e) =>
      e.cellKey === oldCellKey && e.note === oldNote
        ? { ...e, note: newNote }
        : e
    ),
  };
}

/**
 * Remove a note entry
 */
export function removeNote(
  metadata: CellMetadata,
  cellKey: string,
  note: string
): CellMetadata {
  return {
    ...metadata,
    notes: metadata.notes.filter(
      (e) => !(e.cellKey === cellKey && e.note === note)
    ),
  };
}

/**
 * Compare two metadata objects and find differences
 */
export function compareMetadata(
  local: CellMetadata | null,
  excel: CellMetadata | null
): MetadataConflict {
  const differences: MetadataDifference[] = [];

  const localNames = new Map(
    local?.names.map((n) => [n.cellKey, n.name]) || []
  );
  const excelNames = new Map(
    excel?.names.map((n) => [n.cellKey, n.name]) || []
  );

  // Check names
  const allNameKeys = new Set([...localNames.keys(), ...excelNames.keys()]);
  for (const key of allNameKeys) {
    const localVal = localNames.get(key);
    const excelVal = excelNames.get(key);

    if (localVal && !excelVal) {
      differences.push({
        type: "name",
        cellKey: key,
        localValue: localVal,
        status: "only-local",
      });
    } else if (!localVal && excelVal) {
      differences.push({
        type: "name",
        cellKey: key,
        excelValue: excelVal,
        status: "only-excel",
      });
    } else if (localVal !== excelVal) {
      differences.push({
        type: "name",
        cellKey: key,
        localValue: localVal,
        excelValue: excelVal,
        status: "different",
      });
    }
  }

  // Check notes (compare by cellKey + note content)
  const localNotes = new Set(
    local?.notes.map((n) => `${n.cellKey}|||${n.note}`) || []
  );
  const excelNotes = new Set(
    excel?.notes.map((n) => `${n.cellKey}|||${n.note}`) || []
  );

  for (const noteKey of localNotes) {
    if (!excelNotes.has(noteKey)) {
      const [cellKey, note] = noteKey.split("|||");
      differences.push({
        type: "note",
        cellKey,
        localValue: note,
        status: "only-local",
      });
    }
  }

  for (const noteKey of excelNotes) {
    if (!localNotes.has(noteKey)) {
      const [cellKey, note] = noteKey.split("|||");
      differences.push({
        type: "note",
        cellKey,
        excelValue: note,
        status: "only-excel",
      });
    }
  }

  return {
    hasConflict: differences.length > 0,
    localStorage: local,
    excelSheet: excel,
    differences,
  };
}

/**
 * Build a map of cellId -> name for quick lookup during rendering
 */
export function buildCellNameMap(metadata: CellMetadata): Map<string, string> {
  const map = new Map<string, string>();

  for (const entry of metadata.names) {
    const cells = expandCellKey(entry.cellKey);
    for (const cell of cells) {
      map.set(cell, entry.name);
    }
  }

  return map;
}
