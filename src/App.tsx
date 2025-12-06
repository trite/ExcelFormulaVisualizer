import { useState, useCallback, useMemo } from "react";
import {
  ThemeProvider,
  createTheme,
  CssBaseline,
  AppBar,
  Toolbar,
  Typography,
  Container,
  Box,
  Paper,
  IconButton,
  Tooltip,
  Chip,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Alert,
  List,
  ListItem,
  ListItemText,
  Divider,
} from "@mui/material";
import RestartAltIcon from "@mui/icons-material/RestartAlt";
import BubbleChartIcon from "@mui/icons-material/BubbleChart";
import DownloadIcon from "@mui/icons-material/Download";
import DeleteForeverIcon from "@mui/icons-material/DeleteForever";
import GitHubIcon from "@mui/icons-material/GitHub";
import * as XLSX from "xlsx";

import { FileUpload } from "./components/FileUpload";
import { ForceGraphView } from "./components/ForceGraphView";
import { NodeDetails } from "./components/NodeDetails";
import {
  parseExcelFileWithMetadata,
  workbookToGraphData,
  graphDataToForceGraph,
  downloadWorkbookWithMetadata,
} from "./utils/excelParser";
import {
  createEmptyMetadata,
  loadMetadataFromStorage,
  saveMetadataToStorage,
  clearMetadataFromStorage,
  compareMetadata,
  setName,
  removeName,
  addNote,
  updateNote,
  removeNote,
  buildCellNameMap,
} from "./utils/metadataStore";
import {
  filterGraph,
  DEFAULT_FILTER_OPTIONS,
} from "./utils/graphFilter";
import type {
  WorkbookData,
  GraphData,
  GraphNode,
  CellMetadata,
  MetadataConflict,
} from "./types";

const darkTheme = createTheme({
  palette: {
    mode: "dark",
    primary: {
      main: "#90caf9",
    },
    secondary: {
      main: "#f48fb1",
    },
    background: {
      default: "#121212",
      paper: "#1e1e1e",
    },
  },
});

function App() {
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [rawWorkbook, setRawWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [graphData, setGraphData] = useState<GraphData | null>(null);
  const [selectedNode, setSelectedNode] = useState<GraphNode | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Metadata state
  const [cellMetadata, setCellMetadata] = useState<CellMetadata>(
    createEmptyMetadata()
  );
  const [metadataConflict, setMetadataConflict] =
    useState<MetadataConflict | null>(null);
  const [showConflictDialog, setShowConflictDialog] = useState(false);
  const [showClearConfirmDialog, setShowClearConfirmDialog] = useState(false);

  // Filter state
  const [maxNodes, setMaxNodes] = useState(DEFAULT_FILTER_OPTIONS.maxNodes);
  const [showOnlyFormulas, setShowOnlyFormulas] = useState(DEFAULT_FILTER_OPTIONS.showOnlyFormulas);
  const [selectedSheet, setSelectedSheet] = useState<string>("all");

  // Build cell name map for quick lookups
  const cellNameMap = useMemo(
    () => buildCellNameMap(cellMetadata),
    [cellMetadata]
  );

  // Build set of named cell IDs for filtering priority
  const namedCellIds = useMemo(() => {
    const ids = new Set<string>();
    for (const entry of cellMetadata.names) {
      // Cell keys can be comma-separated lists
      const cellKeys = entry.cellKey.split(",").map((k) => k.trim());
      cellKeys.forEach((key) => ids.add(key));
    }
    return ids;
  }, [cellMetadata]);

  // Apply filtering to graph data
  const { filteredGraphData, filterStats } = useMemo(() => {
    if (!graphData) {
      return {
        filteredGraphData: null,
        filterStats: {
          totalNodes: 0,
          totalEdges: 0,
          visibleNodes: 0,
          visibleEdges: 0,
          hiddenBySheet: 0,
          hiddenByFormula: 0,
          hiddenByLimit: 0,
          limitReached: false,
        },
      };
    }

    const result = filterGraph(graphData, {
      maxNodes,
      showOnlyFormulas,
      selectedSheets: selectedSheet === "all" ? "all" : [selectedSheet],
      prioritizeNamed: true,
      namedCells: namedCellIds,
    });

    return {
      filteredGraphData: result.graphData,
      filterStats: result.stats,
    };
  }, [graphData, maxNodes, showOnlyFormulas, selectedSheet, namedCellIds]);

  // Convert filtered graph to force graph format
  const forceGraphData = useMemo(() => {
    if (!filteredGraphData || !workbook) return null;
    return graphDataToForceGraph(filteredGraphData, workbook.sheets);
  }, [filteredGraphData, workbook]);

  const handleFileSelect = useCallback(async (file: File) => {
    setIsLoading(true);
    setError(null);
    try {
      const {
        workbookData,
        excelMetadata,
        rawWorkbook: wb,
      } = await parseExcelFileWithMetadata(file);
      const graph = workbookToGraphData(workbookData);

      setWorkbook(workbookData);
      setRawWorkbook(wb);
      setGraphData(graph);
      setSelectedNode(null);
      // Reset filter to show all sheets when loading new file
      setSelectedSheet("all");

      // Check for metadata conflicts
      const localMetadata = loadMetadataFromStorage(workbookData.fileName);
      const conflict = compareMetadata(localMetadata, excelMetadata);

      if (conflict.hasConflict && localMetadata && excelMetadata) {
        // Both sources have metadata with differences - show dialog
        setMetadataConflict(conflict);
        setShowConflictDialog(true);
        // Temporarily use local metadata until user decides
        setCellMetadata(localMetadata);
      } else if (localMetadata) {
        // Only local exists or no differences
        setCellMetadata(localMetadata);
      } else if (excelMetadata) {
        // Only Excel exists
        setCellMetadata(excelMetadata);
        saveMetadataToStorage(workbookData.fileName, excelMetadata);
      } else {
        // No metadata anywhere
        setCellMetadata(createEmptyMetadata());
      }
    } catch (err) {
      console.error("Error parsing Excel file:", err);
      setError(
        err instanceof Error ? err.message : "Failed to parse Excel file"
      );
    } finally {
      setIsLoading(false);
    }
  }, []);

  const handleResolveConflict = useCallback(
    (useLocal: boolean) => {
      if (!metadataConflict || !workbook) return;

      const chosen = useLocal
        ? metadataConflict.localStorage
        : metadataConflict.excelSheet;
      if (chosen) {
        setCellMetadata(chosen);
        saveMetadataToStorage(workbook.fileName, chosen);
      }

      setShowConflictDialog(false);
      setMetadataConflict(null);
    },
    [metadataConflict, workbook]
  );

  const handleReset = useCallback(() => {
    setWorkbook(null);
    setRawWorkbook(null);
    setGraphData(null);
    setSelectedNode(null);
    setError(null);
    setCellMetadata(createEmptyMetadata());
    setMetadataConflict(null);
    setSelectedSheet("all");
  }, []);

  // Metadata update handlers
  const handleSetName = useCallback(
    (cellKey: string, name: string) => {
      if (!workbook) return;
      const updated = setName(cellMetadata, cellKey, name);
      setCellMetadata(updated);
      saveMetadataToStorage(workbook.fileName, updated);
    },
    [cellMetadata, workbook]
  );

  const handleRemoveName = useCallback(
    (cellKey: string) => {
      if (!workbook) return;
      const updated = removeName(cellMetadata, cellKey);
      setCellMetadata(updated);
      saveMetadataToStorage(workbook.fileName, updated);
    },
    [cellMetadata, workbook]
  );

  const handleAddNote = useCallback(
    (cellKey: string, note: string) => {
      if (!workbook) return;
      const updated = addNote(cellMetadata, cellKey, note);
      setCellMetadata(updated);
      saveMetadataToStorage(workbook.fileName, updated);
    },
    [cellMetadata, workbook]
  );

  const handleUpdateNote = useCallback(
    (cellKey: string, oldNote: string, newNote: string) => {
      if (!workbook) return;
      const updated = updateNote(cellMetadata, cellKey, oldNote, newNote);
      setCellMetadata(updated);
      saveMetadataToStorage(workbook.fileName, updated);
    },
    [cellMetadata, workbook]
  );

  const handleRemoveNote = useCallback(
    (cellKey: string, note: string) => {
      if (!workbook) return;
      const updated = removeNote(cellMetadata, cellKey, note);
      setCellMetadata(updated);
      saveMetadataToStorage(workbook.fileName, updated);
    },
    [cellMetadata, workbook]
  );

  const handleDownloadWithMetadata = useCallback(() => {
    if (!rawWorkbook || !workbook) return;
    downloadWorkbookWithMetadata(rawWorkbook, cellMetadata, workbook.fileName);
  }, [rawWorkbook, workbook, cellMetadata]);

  const handleClearMetadata = useCallback(() => {
    setShowClearConfirmDialog(true);
  }, []);

  const handleConfirmClearMetadata = useCallback(() => {
    if (!workbook) return;
    setCellMetadata(createEmptyMetadata());
    clearMetadataFromStorage(workbook.fileName);
    setShowClearConfirmDialog(false);
  }, [workbook]);

  // Memoize sheets array to prevent unnecessary re-renders of graph components
  const sheets = useMemo(
    () => workbook?.sheets.map((s) => s.name) || [],
    [workbook]
  );

  // Count metadata entries for display
  const metadataCount = cellMetadata.names.length + cellMetadata.notes.length;

  return (
    <ThemeProvider theme={darkTheme}>
      <CssBaseline />
      <Box sx={{ display: "flex", flexDirection: "column", height: "100vh" }}>
        <AppBar position="static" color="default" elevation={1}>
          <Toolbar>
            <BubbleChartIcon sx={{ mr: 2 }} />
            <Typography variant="h6" component="div">
              Excel Formula Visualizer
            </Typography>
            <Box sx={{ flexGrow: 1 }} />
            {workbook && (
              <>
                <Chip
                  label={workbook.fileName}
                  variant="outlined"
                  size="small"
                  sx={{ mr: 2 }}
                />
                <Chip
                  label={`${graphData?.nodes.length || 0} nodes`}
                  size="small"
                  color="primary"
                  sx={{ mr: 2 }}
                />
                <Chip
                  label={`${graphData?.edges.length || 0} edges`}
                  size="small"
                  color="secondary"
                  sx={{ mr: 2 }}
                />
                {metadataCount > 0 && (
                  <Chip
                    label={`${metadataCount} annotations`}
                    size="small"
                    color="info"
                    sx={{ mr: 2 }}
                  />
                )}
                <Tooltip title="Download with metadata">
                  <IconButton
                    onClick={handleDownloadWithMetadata}
                    color="inherit"
                  >
                    <DownloadIcon />
                  </IconButton>
                </Tooltip>
                <Tooltip title="Clear all metadata">
                  <IconButton onClick={handleClearMetadata} color="inherit">
                    <DeleteForeverIcon />
                  </IconButton>
                </Tooltip>
                <Tooltip title="Load new file">
                  <IconButton onClick={handleReset} color="inherit">
                    <RestartAltIcon />
                  </IconButton>
                </Tooltip>
              </>
            )}
            <Tooltip title="View on GitHub">
              <IconButton
                color="inherit"
                href="https://github.com/trite/ExcelFormulaVisualizer"
                target="_blank"
                rel="noopener noreferrer"
              >
                <GitHubIcon />
              </IconButton>
            </Tooltip>
          </Toolbar>
        </AppBar>

        <Box sx={{ flexGrow: 1, overflow: "hidden", p: 2 }}>
          {!workbook ? (
            <Container maxWidth="sm" sx={{ mt: 8 }}>
              <Typography variant="h4" align="center" gutterBottom>
                Visualize Excel Formula Relationships
              </Typography>
              <Typography
                variant="body1"
                align="center"
                color="text.secondary"
                sx={{ mb: 4 }}
              >
                Upload an Excel file to see how formulas reference each other
                across sheets.
              </Typography>
              <FileUpload
                onFileSelect={handleFileSelect}
                isLoading={isLoading}
              />
              {error && (
                <Paper sx={{ mt: 2, p: 2, bgcolor: "error.dark" }}>
                  <Typography color="error.contrastText">{error}</Typography>
                </Paper>
              )}
            </Container>
          ) : (
            <Box sx={{ display: "flex", height: "100%", gap: 2 }}>
              <Paper sx={{ flexGrow: 1, overflow: "hidden" }}>
                {forceGraphData && (
                  <ForceGraphView
                    data={forceGraphData}
                    sheets={sheets}
                    selectedNodeId={selectedNode?.id}
                    onNodeSelect={setSelectedNode}
                    cellNameMap={cellNameMap}
                    selectedSheet={selectedSheet}
                    onSelectedSheetChange={setSelectedSheet}
                    maxNodes={maxNodes}
                    onMaxNodesChange={setMaxNodes}
                    showOnlyFormulas={showOnlyFormulas}
                    onShowOnlyFormulasChange={setShowOnlyFormulas}
                    filterStats={filterStats}
                  />
                )}
              </Paper>
              <Box sx={{ width: 350, flexShrink: 0 }}>
                <NodeDetails
                  node={selectedNode}
                  graphData={graphData}
                  onNodeSelect={setSelectedNode}
                  cellMetadata={cellMetadata}
                  onSetName={handleSetName}
                  onRemoveName={handleRemoveName}
                  onAddNote={handleAddNote}
                  onUpdateNote={handleUpdateNote}
                  onRemoveNote={handleRemoveNote}
                />
              </Box>
            </Box>
          )}
        </Box>
      </Box>

      {/* Metadata Conflict Resolution Dialog */}
      <Dialog
        open={showConflictDialog}
        onClose={() => handleResolveConflict(true)}
        maxWidth="md"
        fullWidth
      >
        <DialogTitle>Metadata Conflict Detected</DialogTitle>
        <DialogContent>
          <Alert severity="warning" sx={{ mb: 2 }}>
            This file has metadata stored in both your browser and the Excel
            file. Please choose which version to keep.
          </Alert>
          <Typography variant="subtitle2" gutterBottom>
            Differences found:
          </Typography>
          <List dense>
            {metadataConflict?.differences.map((diff, idx) => (
              <ListItem key={idx}>
                <ListItemText
                  primary={`${diff.type === "name" ? "Name" : "Note"}: ${
                    diff.cellKey
                  }`}
                  secondary={
                    diff.status === "only-local"
                      ? `Browser only: "${diff.localValue}"`
                      : diff.status === "only-excel"
                      ? `Excel only: "${diff.excelValue}"`
                      : `Browser: "${diff.localValue}" vs Excel: "${diff.excelValue}"`
                  }
                  secondaryTypographyProps={{
                    color:
                      diff.status === "different" ? "error" : "textSecondary",
                  }}
                />
              </ListItem>
            ))}
          </List>
          <Divider sx={{ my: 2 }} />
          <Box sx={{ display: "flex", gap: 2 }}>
            <Box sx={{ flex: 1 }}>
              <Typography variant="subtitle2" color="primary">
                Browser Storage (
                {metadataConflict?.localStorage?.names.length || 0} names,{" "}
                {metadataConflict?.localStorage?.notes.length || 0} notes)
              </Typography>
              <Typography variant="body2" color="text.secondary">
                Your most recent edits in this browser
              </Typography>
            </Box>
            <Box sx={{ flex: 1 }}>
              <Typography variant="subtitle2" color="secondary">
                Excel File ({metadataConflict?.excelSheet?.names.length || 0}{" "}
                names, {metadataConflict?.excelSheet?.notes.length || 0} notes)
              </Typography>
              <Typography variant="body2" color="text.secondary">
                Metadata saved in the Excel file
              </Typography>
            </Box>
          </Box>
        </DialogContent>
        <DialogActions>
          <Button
            onClick={() => handleResolveConflict(false)}
            color="secondary"
          >
            Use Excel File
          </Button>
          <Button
            onClick={() => handleResolveConflict(true)}
            color="primary"
            variant="contained"
          >
            Use Browser Storage
          </Button>
        </DialogActions>
      </Dialog>

      {/* Clear Metadata Confirmation Dialog */}
      <Dialog
        open={showClearConfirmDialog}
        onClose={() => setShowClearConfirmDialog(false)}
      >
        <DialogTitle>Clear All Metadata?</DialogTitle>
        <DialogContent>
          <Alert severity="warning" sx={{ mb: 2 }}>
            This will permanently delete all names and notes for this file from
            your browser.
          </Alert>
          <Typography>
            You have {cellMetadata.names.length} name(s) and{" "}
            {cellMetadata.notes.length} note(s). This action cannot be undone.
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setShowClearConfirmDialog(false)}>
            Cancel
          </Button>
          <Button
            onClick={handleConfirmClearMetadata}
            color="error"
            variant="contained"
          >
            Clear All
          </Button>
        </DialogActions>
      </Dialog>
    </ThemeProvider>
  );
}

export default App;
