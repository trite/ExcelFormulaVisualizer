import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import {
  Box,
  List,
  ListItemButton,
  ListItemIcon,
  ListItemText,
  Collapse,
  Typography,
  Chip,
} from "@mui/material";
import FolderIcon from "@mui/icons-material/Folder";
import FolderOpenIcon from "@mui/icons-material/FolderOpen";
import FunctionsIcon from "@mui/icons-material/Functions";
import GridOnIcon from "@mui/icons-material/GridOn";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import ChevronRightIcon from "@mui/icons-material/ChevronRight";
import type { GraphData, GraphNode } from "../types";

interface CellTreeProps {
  graphData: GraphData | null;
  selectedNodes: Set<string>;
  onSelectionChange: (nodes: Set<string>) => void;
  searchMatchIds?: Set<string> | null;
  cellNameMap?: Map<string, string>;
}

interface SheetGroup {
  name: string;
  cells: GraphNode[];
}

// Sort cell addresses naturally (A1, A2, A10, B1, etc.)
function compareCellAddresses(a: string, b: string): number {
  const parseAddress = (addr: string) => {
    const match = addr.match(/^([A-Za-z]+)(\d+)$/);
    if (!match) return { col: addr, row: 0 };
    return { col: match[1].toUpperCase(), row: parseInt(match[2], 10) };
  };

  const addrA = parseAddress(a);
  const addrB = parseAddress(b);

  // Compare columns first (alphabetically)
  if (addrA.col !== addrB.col) {
    return addrA.col.localeCompare(addrB.col);
  }
  // Then compare rows (numerically)
  return addrA.row - addrB.row;
}

export function CellTree({
  graphData,
  selectedNodes,
  onSelectionChange,
  searchMatchIds,
  cellNameMap,
}: CellTreeProps) {
  const [expandedSheets, setExpandedSheets] = useState<Set<string>>(new Set());
  const lastClickedRef = useRef<string | null>(null);
  const itemRefs = useRef<Map<string, HTMLDivElement>>(new Map());

  // Group nodes by sheet
  const sheetGroups = useMemo((): SheetGroup[] => {
    if (!graphData) return [];

    const groups = new Map<string, GraphNode[]>();

    for (const node of graphData.nodes) {
      // If search is active, only include matching nodes
      if (searchMatchIds && !searchMatchIds.has(node.id)) {
        continue;
      }

      if (!groups.has(node.sheet)) {
        groups.set(node.sheet, []);
      }
      groups.get(node.sheet)!.push(node);
    }

    // Sort cells within each group and create result
    const result: SheetGroup[] = [];
    for (const [name, cells] of groups) {
      cells.sort((a, b) => compareCellAddresses(a.address, b.address));
      result.push({ name, cells });
    }

    // Sort sheets alphabetically
    result.sort((a, b) => a.name.localeCompare(b.name));

    return result;
  }, [graphData, searchMatchIds]);

  // Build flat list of visible cell IDs for range selection
  const visibleCellIds = useMemo(() => {
    const ids: string[] = [];
    for (const group of sheetGroups) {
      if (expandedSheets.has(group.name)) {
        for (const cell of group.cells) {
          ids.push(cell.id);
        }
      }
    }
    return ids;
  }, [sheetGroups, expandedSheets]);

  // Auto-expand sheets that contain selected nodes
  useEffect(() => {
    if (selectedNodes.size > 0) {
      const sheetsToExpand = new Set<string>();
      for (const nodeId of selectedNodes) {
        // Extract sheet name from node ID (format: "Sheet1!A1")
        const match = nodeId.match(/^(.+)!/);
        if (match) {
          sheetsToExpand.add(match[1]);
        }
      }
      if (sheetsToExpand.size > 0) {
        setExpandedSheets((prev) => {
          const next = new Set(prev);
          for (const sheet of sheetsToExpand) {
            next.add(sheet);
          }
          return next;
        });
      }
    }
  }, [selectedNodes]);

  // Auto-scroll to first selected node
  useEffect(() => {
    if (selectedNodes.size > 0) {
      const firstSelectedId = Array.from(selectedNodes)[0];
      const ref = itemRefs.current.get(firstSelectedId);
      if (ref) {
        ref.scrollIntoView({ behavior: "smooth", block: "nearest" });
      }
    }
  }, [selectedNodes]);

  const toggleSheet = useCallback((sheetName: string) => {
    setExpandedSheets((prev) => {
      const next = new Set(prev);
      if (next.has(sheetName)) {
        next.delete(sheetName);
      } else {
        next.add(sheetName);
      }
      return next;
    });
  }, []);

  const handleCellClick = useCallback(
    (node: GraphNode, event: React.MouseEvent) => {
      event.stopPropagation();

      if (event.shiftKey && lastClickedRef.current) {
        // Range selection
        const lastIndex = visibleCellIds.indexOf(lastClickedRef.current);
        const currentIndex = visibleCellIds.indexOf(node.id);

        if (lastIndex !== -1 && currentIndex !== -1) {
          const start = Math.min(lastIndex, currentIndex);
          const end = Math.max(lastIndex, currentIndex);
          const rangeIds = visibleCellIds.slice(start, end + 1);

          if (event.ctrlKey || event.metaKey) {
            // Add range to existing selection
            const newSelection = new Set(selectedNodes);
            rangeIds.forEach((id) => newSelection.add(id));
            onSelectionChange(newSelection);
          } else {
            // Replace with range
            onSelectionChange(new Set(rangeIds));
          }
        }
      } else if (event.ctrlKey || event.metaKey) {
        // Toggle individual node
        const newSelection = new Set(selectedNodes);
        if (newSelection.has(node.id)) {
          newSelection.delete(node.id);
        } else {
          newSelection.add(node.id);
        }
        onSelectionChange(newSelection);
        lastClickedRef.current = node.id;
      } else {
        // Single selection
        onSelectionChange(new Set([node.id]));
        lastClickedRef.current = node.id;
      }
    },
    [selectedNodes, onSelectionChange, visibleCellIds]
  );

  const formatValue = (value: unknown): string => {
    if (value === undefined || value === null) return "";
    if (typeof value === "string") {
      return value.length > 20 ? value.slice(0, 20) + "..." : value;
    }
    if (typeof value === "number") {
      return value.toLocaleString();
    }
    return String(value);
  };

  if (!graphData || graphData.nodes.length === 0) {
    return (
      <Box sx={{ p: 2 }}>
        <Typography variant="body2" color="text.secondary">
          No cells to display
        </Typography>
      </Box>
    );
  }

  if (searchMatchIds && searchMatchIds.size === 0) {
    return (
      <Box sx={{ p: 2 }}>
        <Typography variant="body2" color="text.secondary">
          No matching cells
        </Typography>
      </Box>
    );
  }

  return (
    <List dense disablePadding sx={{ overflow: "auto" }}>
      {sheetGroups.map((group) => {
        const isExpanded = expandedSheets.has(group.name);
        const hasSelectedCells = group.cells.some((c) =>
          selectedNodes.has(c.id)
        );

        return (
          <Box key={group.name}>
            {/* Sheet folder header */}
            <ListItemButton
              onClick={() => toggleSheet(group.name)}
              sx={{
                bgcolor: hasSelectedCells ? "action.selected" : undefined,
              }}
            >
              <ListItemIcon sx={{ minWidth: 32 }}>
                {isExpanded ? (
                  <FolderOpenIcon fontSize="small" color="primary" />
                ) : (
                  <FolderIcon fontSize="small" color="primary" />
                )}
              </ListItemIcon>
              <ListItemText
                primary={group.name}
                secondary={`${group.cells.length} cells`}
                slotProps={{
                  primary: { variant: "body2", fontWeight: 500 },
                  secondary: { variant: "caption" },
                }}
              />
              {isExpanded ? (
                <ExpandMoreIcon fontSize="small" />
              ) : (
                <ChevronRightIcon fontSize="small" />
              )}
            </ListItemButton>

            {/* Cells under this sheet */}
            <Collapse in={isExpanded} timeout="auto" unmountOnExit>
              <List dense disablePadding>
                {group.cells.map((cell) => {
                  const isSelected = selectedNodes.has(cell.id);
                  const cellName = cellNameMap?.get(cell.id);
                  const displayName = cellName || cell.excelName;

                  return (
                    <ListItemButton
                      key={cell.id}
                      ref={(el) => {
                        if (el) {
                          itemRefs.current.set(cell.id, el);
                        }
                      }}
                      selected={isSelected}
                      onClick={(e) => handleCellClick(cell, e)}
                      sx={{ pl: 4 }}
                    >
                      <ListItemIcon sx={{ minWidth: 28 }}>
                        {cell.hasFormula ? (
                          <FunctionsIcon
                            sx={{ fontSize: 16 }}
                            color="primary"
                          />
                        ) : (
                          <GridOnIcon sx={{ fontSize: 16 }} color="action" />
                        )}
                      </ListItemIcon>
                      <ListItemText
                        primary={
                          <Box
                            sx={{
                              display: "flex",
                              alignItems: "center",
                              gap: 0.5,
                            }}
                          >
                            <Typography variant="body2">
                              {cell.address}
                            </Typography>
                            {displayName && (
                              <Chip
                                label={displayName}
                                size="small"
                                sx={{
                                  height: 18,
                                  fontSize: "0.65rem",
                                  "& .MuiChip-label": { px: 0.75 },
                                }}
                                color={cellName ? "info" : "secondary"}
                                variant={cellName ? "filled" : "outlined"}
                              />
                            )}
                          </Box>
                        }
                        secondary={
                          cell.hasFormula
                            ? cell.formula?.slice(0, 30) +
                              (cell.formula && cell.formula.length > 30
                                ? "..."
                                : "")
                            : formatValue(cell.value)
                        }
                        slotProps={{
                          secondary: {
                            variant: "caption",
                            sx: {
                              fontFamily: cell.hasFormula
                                ? "monospace"
                                : undefined,
                              fontSize: "0.7rem",
                            },
                          },
                        }}
                      />
                    </ListItemButton>
                  );
                })}
              </List>
            </Collapse>
          </Box>
        );
      })}
    </List>
  );
}
