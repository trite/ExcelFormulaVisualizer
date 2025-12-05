// Formula parser to extract cell references from Excel formulas

/**
 * Extracts all cell references from an Excel formula
 * Handles:
 * - Simple references: A1, B2, $A$1, A$1, $A1
 * - Ranges: A1:B5, $A$1:$B$5
 * - Cross-sheet references: Sheet1!A1, 'Sheet Name'!A1
 * - Cross-sheet ranges: Sheet1!A1:B5
 */
export function extractCellReferences(
  formula: string,
  currentSheet: string
): string[] {
  const references: Set<string> = new Set();

  if (!formula || !formula.startsWith("=")) {
    return [];
  }

  // Pattern for sheet name (handles quoted names with spaces)
  const sheetPattern = "(?:'([^']+)'|([A-Za-z0-9_]+))!";

  // Pattern for cell reference (with optional $ for absolute references)
  const cellPattern = "\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7}";

  // Combined pattern for cross-sheet reference with optional range
  const crossSheetRangeRegex = new RegExp(
    `${sheetPattern}(${cellPattern}(?::${cellPattern})?)`,
    "g"
  );

  // Pattern for simple cell reference or range (no sheet prefix)
  const simpleCellRegex = new RegExp(
    `(?<![A-Za-z0-9_!')])((${cellPattern})(?::(${cellPattern}))?)(?![A-Za-z0-9(])`,
    "g"
  );

  // First, find all cross-sheet references
  let match;
  const processedRanges: string[] = [];

  while ((match = crossSheetRangeRegex.exec(formula)) !== null) {
    const sheetName = match[1] || match[2]; // quoted or unquoted sheet name
    const cellOrRange = match[3];
    processedRanges.push(match[0]);

    if (cellOrRange.includes(":")) {
      // It's a range, expand it
      const [startCell, endCell] = cellOrRange.split(":");
      const expandedCells = expandRange(
        normalizeCell(startCell),
        normalizeCell(endCell)
      );
      expandedCells.forEach((cell) => references.add(`${sheetName}!${cell}`));
    } else {
      references.add(`${sheetName}!${normalizeCell(cellOrRange)}`);
    }
  }

  // Remove cross-sheet references from formula to avoid double-matching
  let formulaWithoutCrossSheet = formula;
  processedRanges.forEach((range) => {
    formulaWithoutCrossSheet = formulaWithoutCrossSheet.replace(range, " ");
  });

  // Now find simple references (same sheet)
  while ((match = simpleCellRegex.exec(formulaWithoutCrossSheet)) !== null) {
    const cellOrRange = match[1];

    // Skip if it looks like it's part of a function name or other identifier
    if (/^[A-Za-z]+$/.test(cellOrRange)) {
      continue;
    }

    if (cellOrRange.includes(":")) {
      // It's a range, expand it
      const [startCell, endCell] = cellOrRange.split(":");
      const expandedCells = expandRange(
        normalizeCell(startCell),
        normalizeCell(endCell)
      );
      expandedCells.forEach((cell) =>
        references.add(`${currentSheet}!${cell}`)
      );
    } else {
      references.add(`${currentSheet}!${normalizeCell(cellOrRange)}`);
    }
  }

  return Array.from(references);
}

/**
 * Removes $ signs from cell references to normalize them
 */
function normalizeCell(cell: string): string {
  return cell.replace(/\$/g, "").toUpperCase();
}

/**
 * Expands a range like A1:B3 into individual cells
 */
function expandRange(startCell: string, endCell: string): string[] {
  const cells: string[] = [];

  const startCol = getColumnIndex(startCell);
  const startRow = getRowNumber(startCell);
  const endCol = getColumnIndex(endCell);
  const endRow = getRowNumber(endCell);

  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);
  const minRow = Math.min(startRow, endRow);
  const maxRow = Math.max(startRow, endRow);

  // Limit expansion to prevent memory issues with huge ranges
  const maxCells = 1000;
  const totalCells = (maxCol - minCol + 1) * (maxRow - minRow + 1);

  if (totalCells > maxCells) {
    // For very large ranges, just return corner cells to indicate the relationship
    cells.push(getColumnLetter(minCol) + minRow);
    cells.push(getColumnLetter(maxCol) + maxRow);
    return cells;
  }

  for (let col = minCol; col <= maxCol; col++) {
    for (let row = minRow; row <= maxRow; row++) {
      cells.push(getColumnLetter(col) + row);
    }
  }

  return cells;
}

/**
 * Converts column letter(s) to 0-based index
 */
function getColumnIndex(cell: string): number {
  const match = cell.match(/^([A-Z]+)/i);
  if (!match) return 0;

  const letters = match[1].toUpperCase();
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index;
}

/**
 * Converts 1-based column index to letter(s)
 */
function getColumnLetter(index: number): string {
  let letter = "";
  while (index > 0) {
    const remainder = (index - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

/**
 * Extracts row number from cell reference
 */
function getRowNumber(cell: string): number {
  const match = cell.match(/([0-9]+)$/);
  return match ? parseInt(match[1], 10) : 1;
}
