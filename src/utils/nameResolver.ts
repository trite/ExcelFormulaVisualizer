/**
 * Name resolver for Excel defined names (from Name Manager)
 * Handles resolution of named references to cell addresses and
 * provides display names for cells that have Excel-defined names.
 */

import type { DefinedNameInfo } from "../types";

/**
 * Represents a resolved name with its cell references
 */
export interface ResolvedName {
  name: string;
  displayName: string; // May include scope prefix if ambiguous
  cells: string[]; // Full addresses (e.g., "Sheet1!A1")
  isRange: boolean;
  sheetScope?: number;
}

/**
 * Map from name (possibly with scope) to resolved cell addresses
 * - Global names: key is just the name (e.g., "mass")
 * - Sheet-scoped names: key is "SheetIndex::name" (e.g., "0::mass")
 */
export class NameResolver {
  // Maps name -> ResolvedName (for global names)
  private globalNames = new Map<string, ResolvedName>();

  // Maps "sheetIndex::name" -> ResolvedName (for sheet-scoped names)
  private scopedNames = new Map<string, ResolvedName>();

  // Maps cell fullAddress -> display name (for reverse lookup)
  private cellToName = new Map<string, string>();

  // Track which names are ambiguous (exist at multiple scopes)
  private ambiguousNames = new Set<string>();

  // Sheet names for resolving scope
  private sheetNames: string[] = [];

  constructor(definedNames: DefinedNameInfo[], sheetNames: string[]) {
    this.sheetNames = sheetNames;
    this.buildMaps(definedNames);
  }

  private buildMaps(definedNames: DefinedNameInfo[]): void {
    // First pass: identify ambiguous names
    const nameOccurrences = new Map<string, number>();
    for (const def of definedNames) {
      const count = nameOccurrences.get(def.name) || 0;
      nameOccurrences.set(def.name, count + 1);
    }
    for (const [name, count] of nameOccurrences) {
      if (count > 1) {
        this.ambiguousNames.add(name);
      }
    }

    // Second pass: resolve names to cells
    for (const def of definedNames) {
      const resolved = this.resolveDefinition(def);
      if (!resolved) continue;

      if (def.sheetScope !== undefined) {
        // Sheet-scoped name
        const scopeKey = `${def.sheetScope}::${def.name}`;
        this.scopedNames.set(scopeKey, resolved);
      } else {
        // Global name
        this.globalNames.set(def.name, resolved);
      }

      // Build reverse mapping (cell -> name) for display purposes
      // Only map single cells, not ranges (ranges would be confusing)
      if (!resolved.isRange && resolved.cells.length === 1) {
        const existingName = this.cellToName.get(resolved.cells[0]);
        // Prefer global names over scoped names for display
        if (!existingName || def.sheetScope === undefined) {
          this.cellToName.set(resolved.cells[0], resolved.displayName);
        }
      }
    }
  }

  private resolveDefinition(def: DefinedNameInfo): ResolvedName | null {
    const ref = def.ref;

    // Skip invalid or complex references (formulas, errors, etc.)
    if (!ref || ref.startsWith("#") || ref.includes(",")) {
      return null;
    }

    // Parse the reference
    const cells = this.parseReference(ref);
    if (cells.length === 0) return null;

    // Determine display name based on ambiguity
    let displayName = def.name;
    if (this.ambiguousNames.has(def.name) && def.sheetScope !== undefined) {
      const sheetName = this.sheetNames[def.sheetScope] || `Sheet${def.sheetScope + 1}`;
      displayName = `${sheetName}::${def.name}`;
    }

    return {
      name: def.name,
      displayName,
      cells,
      isRange: ref.includes(":"),
      sheetScope: def.sheetScope,
    };
  }

  /**
   * Parse an Excel reference string into full cell addresses
   * Handles: "Sheet1!$A$1", "Sheet1!$A$1:$B$5", "'Sheet Name'!A1"
   */
  private parseReference(ref: string): string[] {
    const cells: string[] = [];

    // Extract sheet name and cell reference
    let sheetName: string;
    let cellRef: string;

    const sheetMatch = ref.match(/^(?:'([^']+)'|([A-Za-z0-9_]+))!(.+)$/);
    if (sheetMatch) {
      sheetName = sheetMatch[1] || sheetMatch[2];
      cellRef = sheetMatch[3];
    } else {
      // No sheet specified - this shouldn't happen for defined names
      return [];
    }

    // Normalize cell reference (remove $ signs)
    cellRef = cellRef.replace(/\$/g, "").toUpperCase();

    if (cellRef.includes(":")) {
      // It's a range - expand it
      const [start, end] = cellRef.split(":");
      const expandedCells = this.expandRange(start, end);
      for (const cell of expandedCells) {
        cells.push(`${sheetName}!${cell}`);
      }
    } else {
      // Single cell
      cells.push(`${sheetName}!${cellRef}`);
    }

    return cells;
  }

  /**
   * Expand a range like A1:B3 into individual cells
   */
  private expandRange(startCell: string, endCell: string): string[] {
    const cells: string[] = [];

    const startCol = this.getColumnIndex(startCell);
    const startRow = this.getRowNumber(startCell);
    const endCol = this.getColumnIndex(endCell);
    const endRow = this.getRowNumber(endCell);

    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);

    // Limit expansion to prevent memory issues
    const maxCells = 1000;
    const totalCells = (maxCol - minCol + 1) * (maxRow - minRow + 1);

    if (totalCells > maxCells) {
      // For very large ranges, just return corner cells
      cells.push(this.getColumnLetter(minCol) + minRow);
      cells.push(this.getColumnLetter(maxCol) + maxRow);
      return cells;
    }

    for (let col = minCol; col <= maxCol; col++) {
      for (let row = minRow; row <= maxRow; row++) {
        cells.push(this.getColumnLetter(col) + row);
      }
    }

    return cells;
  }

  private getColumnIndex(cell: string): number {
    const match = cell.match(/^([A-Z]+)/i);
    if (!match) return 0;
    const letters = match[1].toUpperCase();
    let index = 0;
    for (let i = 0; i < letters.length; i++) {
      index = index * 26 + (letters.charCodeAt(i) - 64);
    }
    return index;
  }

  private getColumnLetter(index: number): string {
    let letter = "";
    while (index > 0) {
      const remainder = (index - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      index = Math.floor((index - 1) / 26);
    }
    return letter;
  }

  private getRowNumber(cell: string): number {
    const match = cell.match(/([0-9]+)$/);
    return match ? parseInt(match[1], 10) : 1;
  }

  /**
   * Resolve a name reference in a formula to cell addresses
   * @param name The name to resolve (e.g., "mass")
   * @param currentSheetIndex The index of the sheet where the formula is (for scoped names)
   * @returns Array of full cell addresses, or empty if not found
   */
  resolveName(name: string, currentSheetIndex: number): string[] {
    // First check for sheet-scoped name on the current sheet
    const scopeKey = `${currentSheetIndex}::${name}`;
    const scopedResolved = this.scopedNames.get(scopeKey);
    if (scopedResolved) {
      return scopedResolved.cells;
    }

    // Then check for global name
    const globalResolved = this.globalNames.get(name);
    if (globalResolved) {
      return globalResolved.cells;
    }

    return [];
  }

  /**
   * Get the display name for a cell (if it has an Excel-defined name)
   * @param fullAddress The full cell address (e.g., "Sheet1!A1")
   * @returns The display name or undefined
   */
  getDisplayName(fullAddress: string): string | undefined {
    return this.cellToName.get(fullAddress);
  }

  /**
   * Get all defined names (for formula parsing)
   * Returns names sorted by length descending to match longer names first
   */
  getAllNames(): string[] {
    const names = new Set<string>();
    for (const [name] of this.globalNames) {
      names.add(name);
    }
    for (const resolved of this.scopedNames.values()) {
      names.add(resolved.name);
    }
    // Sort by length descending so we match longer names first
    return Array.from(names).sort((a, b) => b.length - a.length);
  }

  /**
   * Check if a token could be a defined name
   */
  isDefinedName(token: string): boolean {
    if (this.globalNames.has(token)) return true;
    for (const resolved of this.scopedNames.values()) {
      if (resolved.name === token) return true;
    }
    return false;
  }
}

/**
 * Create a NameResolver from workbook data
 */
export function createNameResolver(
  definedNames: DefinedNameInfo[],
  sheetNames: string[]
): NameResolver {
  return new NameResolver(definedNames, sheetNames);
}
