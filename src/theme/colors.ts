/**
 * Centralized color definitions for the Excel Formula Visualizer
 * Used by both the application and tests for consistency
 */

// Sheet colors - used to color nodes by their sheet
// Applied in order based on sheet index
export const SHEET_COLORS = [
  "#4fc3f7", // cyan
  "#81c784", // green
  "#ffb74d", // orange
  "#f06292", // pink
  "#ba68c8", // purple
  "#4db6ac", // teal
  "#aed581", // light green
  "#ff8a65", // deep orange
] as const;

// Helper to get sheet color by index
export function getSheetColor(index: number): string {
  return SHEET_COLORS[index % SHEET_COLORS.length];
}

// Selection and highlighting colors
export const SELECTION_COLORS = {
  // Glow color for selected nodes
  selected: "#f0f4ff",
  // Color for input nodes (nodes that feed into the selected node)
  input: "#e91e63",
  // Color for output nodes (nodes that the selected node feeds into)
  output: "#cddc39",
} as const;

// Node label colors
export const LABEL_COLORS = {
  // Color for nodes with custom names
  named: "#ffeb3b",
  // Default label color
  default: "#ffffff",
  // Dimmed label color when not selected/connected
  dimmed: "rgba(255,255,255,0.3)",
} as const;

// Link/edge colors
export const LINK_COLORS = {
  // Default link color
  default: "rgba(255,255,255,0.3)",
  // Dimmed link color when not connected to selected node
  dimmed: "rgba(255,255,255,0.08)",
} as const;

// MUI theme palette colors (for reference)
export const THEME_COLORS = {
  primary: "#90caf9",
  secondary: "#f48fb1",
  background: {
    default: "#121212",
    paper: "#1e1e1e",
  },
} as const;
