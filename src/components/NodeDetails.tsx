import { useState, useMemo } from "react";
import {
  Paper,
  Typography,
  Box,
  Chip,
  Divider,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  List,
  ListItemButton,
  ListItemText,
  ListItemIcon,
  TextField,
  Button,
  IconButton,
  Alert,
  Tooltip,
} from "@mui/material";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import FunctionsIcon from "@mui/icons-material/Functions";
import GridOnIcon from "@mui/icons-material/GridOn";
import InputIcon from "@mui/icons-material/Input";
import OutputIcon from "@mui/icons-material/Output";
import LabelIcon from "@mui/icons-material/Label";
import NotesIcon from "@mui/icons-material/Notes";
import DeleteIcon from "@mui/icons-material/Delete";
import EditIcon from "@mui/icons-material/Edit";
import SaveIcon from "@mui/icons-material/Save";
import CancelIcon from "@mui/icons-material/Cancel";
import AddIcon from "@mui/icons-material/Add";
import type {
  GraphNode,
  GraphData,
  CellMetadata,
  CellNoteEntry,
} from "../types";
import {
  findExistingName,
  findNotesForCell,
  validateNameAssignment,
  expandCellKey,
} from "../utils/metadataStore";

interface NodeDetailsProps {
  node: GraphNode | null;
  graphData: GraphData | null;
  onNodeSelect: (node: GraphNode | null) => void;
  cellMetadata: CellMetadata;
  onSetName: (cellKey: string, name: string) => void;
  onRemoveName: (cellKey: string) => void;
  onAddNote: (cellKey: string, note: string) => void;
  onUpdateNote: (cellKey: string, oldNote: string, newNote: string) => void;
  onRemoveNote: (cellKey: string, note: string) => void;
}

interface RelatedNodeInfo {
  node: GraphNode;
  inputs: GraphNode[];
}

export function NodeDetails({
  node,
  graphData,
  onNodeSelect,
  cellMetadata,
  onSetName,
  onRemoveName,
  onAddNote,
  onUpdateNote,
  onRemoveNote,
}: NodeDetailsProps) {
  const [expandedInputs, setExpandedInputs] = useState<Set<string>>(new Set());
  const [expandedOutputs, setExpandedOutputs] = useState<Set<string>>(
    new Set()
  );

  // Name editing state
  const [isEditingName, setIsEditingName] = useState(false);
  const [nameValue, setNameValue] = useState("");
  const [nameError, setNameError] = useState<string | null>(null);

  // Note editing state
  const [isAddingNote, setIsAddingNote] = useState(false);
  const [newNoteValue, setNewNoteValue] = useState("");
  const [editingNote, setEditingNote] = useState<{
    cellKey: string;
    oldNote: string;
  } | null>(null);
  const [editNoteValue, setEditNoteValue] = useState("");

  // Get current name for this cell
  const currentNameEntry = useMemo(() => {
    if (!node) return null;
    return findExistingName(node.id, cellMetadata);
  }, [node, cellMetadata]);

  // Get notes for this cell
  const cellNotes = useMemo(() => {
    if (!node) return [];
    return findNotesForCell(node.id, cellMetadata);
  }, [node, cellMetadata]);

  // Compute input nodes (nodes that this node references)
  const inputNodes = useMemo((): RelatedNodeInfo[] => {
    if (!node || !graphData) return [];

    const inputEdges = graphData.edges.filter((e) => e.target === node.id);
    const inputNodeIds = inputEdges.map((e) => e.source);

    return inputNodeIds
      .map((id) => {
        const inputNode = graphData.nodes.find((n) => n.id === id);
        if (!inputNode) return null;

        const nestedInputEdges = graphData.edges.filter((e) => e.target === id);
        const nestedInputs = nestedInputEdges
          .map((e) => graphData.nodes.find((n) => n.id === e.source))
          .filter((n): n is GraphNode => n !== undefined);

        return { node: inputNode, inputs: nestedInputs };
      })
      .filter((n): n is RelatedNodeInfo => n !== null);
  }, [node, graphData]);

  // Compute output nodes (nodes that reference this node)
  const outputNodes = useMemo((): RelatedNodeInfo[] => {
    if (!node || !graphData) return [];

    const outputEdges = graphData.edges.filter((e) => e.source === node.id);
    const outputNodeIds = outputEdges.map((e) => e.target);

    return outputNodeIds
      .map((id) => {
        const outputNode = graphData.nodes.find((n) => n.id === id);
        if (!outputNode) return null;

        const nestedInputEdges = graphData.edges.filter((e) => e.target === id);
        const nestedInputs = nestedInputEdges
          .map((e) => graphData.nodes.find((n) => n.id === e.source))
          .filter((n): n is GraphNode => n !== undefined);

        return { node: outputNode, inputs: nestedInputs };
      })
      .filter((n): n is RelatedNodeInfo => n !== null);
  }, [node, graphData]);

  const toggleInputExpanded = (nodeId: string) => {
    setExpandedInputs((prev) => {
      const next = new Set(prev);
      if (next.has(nodeId)) {
        next.delete(nodeId);
      } else {
        next.add(nodeId);
      }
      return next;
    });
  };

  const toggleOutputExpanded = (nodeId: string) => {
    setExpandedOutputs((prev) => {
      const next = new Set(prev);
      if (next.has(nodeId)) {
        next.delete(nodeId);
      } else {
        next.add(nodeId);
      }
      return next;
    });
  };

  const formatValue = (value: unknown): string => {
    if (value === undefined || value === null) return "-";
    if (typeof value === "object") return JSON.stringify(value);
    return String(value);
  };

  // Name handlers
  const handleStartEditName = () => {
    setNameValue(currentNameEntry?.name || "");
    setNameError(null);
    setIsEditingName(true);
  };

  const handleSaveName = () => {
    if (!node || !nameValue.trim()) {
      setNameError("Name cannot be empty");
      return;
    }

    // Validate no conflicts (excluding current entry if editing)
    const conflicts = validateNameAssignment(
      node.id,
      cellMetadata,
      currentNameEntry?.cellKey
    );

    if (conflicts.length > 0) {
      setNameError(conflicts[0]);
      return;
    }

    onSetName(node.id, nameValue.trim());
    setIsEditingName(false);
    setNameError(null);
  };

  const handleCancelEditName = () => {
    setIsEditingName(false);
    setNameError(null);
  };

  const handleRemoveName = () => {
    if (currentNameEntry) {
      onRemoveName(currentNameEntry.cellKey);
    }
  };

  // Note handlers
  const handleStartAddNote = () => {
    setNewNoteValue("");
    setIsAddingNote(true);
  };

  const handleSaveNewNote = () => {
    if (!node || !newNoteValue.trim()) return;
    onAddNote(node.id, newNoteValue.trim());
    setIsAddingNote(false);
    setNewNoteValue("");
  };

  const handleCancelAddNote = () => {
    setIsAddingNote(false);
    setNewNoteValue("");
  };

  const handleStartEditNote = (entry: CellNoteEntry) => {
    setEditingNote({ cellKey: entry.cellKey, oldNote: entry.note });
    setEditNoteValue(entry.note);
  };

  const handleSaveEditNote = () => {
    if (!editingNote || !editNoteValue.trim()) return;
    onUpdateNote(
      editingNote.cellKey,
      editingNote.oldNote,
      editNoteValue.trim()
    );
    setEditingNote(null);
    setEditNoteValue("");
  };

  const handleCancelEditNote = () => {
    setEditingNote(null);
    setEditNoteValue("");
  };

  const handleRemoveNote = (entry: CellNoteEntry) => {
    onRemoveNote(entry.cellKey, entry.note);
  };

  // Get other cells affected by a note
  const getOtherCellsForNote = (entry: CellNoteEntry): string[] => {
    if (!node) return [];
    const allCells = expandCellKey(entry.cellKey);
    return allCells.filter((c) => c !== node.id);
  };

  const renderRelatedNode = (
    info: RelatedNodeInfo,
    isExpanded: boolean,
    onToggle: () => void
  ) => (
    <Box key={info.node.id} sx={{ mb: 1 }}>
      <ListItemButton
        onClick={() => onNodeSelect(info.node)}
        sx={{
          borderRadius: 1,
          bgcolor: "action.hover",
          "&:hover": { bgcolor: "action.selected" },
        }}
      >
        <ListItemIcon sx={{ minWidth: 32 }}>
          {info.node.hasFormula ? (
            <FunctionsIcon fontSize="small" color="primary" />
          ) : (
            <GridOnIcon fontSize="small" color="action" />
          )}
        </ListItemIcon>
        <ListItemText
          primary={info.node.id}
          secondary={formatValue(info.node.value)}
          slotProps={{
            primary: { variant: "body2", fontWeight: 500 },
            secondary: { variant: "caption" },
          }}
        />
        {info.inputs.length > 0 && (
          <Box
            component="span"
            onClick={(e) => {
              e.stopPropagation();
              onToggle();
            }}
            sx={{
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              color: "text.secondary",
              "&:hover": { color: "primary.main" },
            }}
          >
            <Typography variant="caption" sx={{ mr: 0.5 }}>
              {info.inputs.length}
            </Typography>
            <ExpandMoreIcon
              fontSize="small"
              sx={{
                transform: isExpanded ? "rotate(180deg)" : "rotate(0deg)",
                transition: "transform 0.2s",
              }}
            />
          </Box>
        )}
      </ListItemButton>

      {isExpanded && info.inputs.length > 0 && (
        <Box sx={{ pl: 3, pt: 1 }}>
          <Typography
            variant="caption"
            color="text.secondary"
            sx={{ mb: 0.5, display: "block" }}
          >
            Inputs to {info.node.address}:
          </Typography>
          {info.node.formula && (
            <Paper
              variant="outlined"
              sx={{
                p: 0.5,
                mb: 1,
                bgcolor: "grey.900",
                fontFamily: "monospace",
                fontSize: "0.75rem",
                wordBreak: "break-all",
              }}
            >
              {info.node.formula}
            </Paper>
          )}
          <List dense disablePadding>
            {info.inputs.map((input) => (
              <ListItemButton
                key={input.id}
                onClick={() => onNodeSelect(input)}
                sx={{
                  py: 0.25,
                  borderRadius: 0.5,
                  "&:hover": { bgcolor: "action.selected" },
                }}
              >
                <ListItemIcon sx={{ minWidth: 24 }}>
                  {input.hasFormula ? (
                    <FunctionsIcon sx={{ fontSize: 14 }} color="primary" />
                  ) : (
                    <GridOnIcon sx={{ fontSize: 14 }} color="action" />
                  )}
                </ListItemIcon>
                <ListItemText
                  primary={input.address}
                  secondary={formatValue(input.value)}
                  slotProps={{
                    primary: { variant: "caption" },
                    secondary: { variant: "caption", fontSize: "0.65rem" },
                  }}
                />
              </ListItemButton>
            ))}
          </List>
        </Box>
      )}
    </Box>
  );

  if (!node) {
    return (
      <Paper sx={{ p: 2, height: "100%" }}>
        <Typography color="text.secondary" sx={{ fontStyle: "italic" }}>
          Click on a node to see its details
        </Typography>
      </Paper>
    );
  }

  return (
    <Paper sx={{ p: 2, height: "100%", overflow: "auto" }}>
      <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 2 }}>
        {node.hasFormula ? (
          <FunctionsIcon color="primary" />
        ) : (
          <GridOnIcon color="action" />
        )}
        <Typography variant="h6">{node.id}</Typography>
      </Box>

      <Divider sx={{ mb: 2 }} />

      <Box sx={{ mb: 2 }}>
        <Typography variant="subtitle2" color="text.secondary">
          Sheet
        </Typography>
        <Chip
          label={node.sheet}
          size="small"
          color="primary"
          variant="outlined"
        />
      </Box>

      <Box sx={{ mb: 2 }}>
        <Typography variant="subtitle2" color="text.secondary">
          Cell Address
        </Typography>
        <Typography variant="body1">{node.address}</Typography>
      </Box>

      {/* Cell Name Section */}
      <Box sx={{ mb: 2 }}>
        <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 1 }}>
          <LabelIcon fontSize="small" color="info" />
          <Typography variant="subtitle2" color="text.secondary">
            Name
          </Typography>
        </Box>

        {isEditingName ? (
          <Box>
            <TextField
              size="small"
              fullWidth
              value={nameValue}
              onChange={(e) => setNameValue(e.target.value)}
              placeholder="Enter a name for this cell"
              error={!!nameError}
              helperText={nameError}
              autoFocus
              sx={{ mb: 1 }}
            />
            <Box sx={{ display: "flex", gap: 1 }}>
              <Button
                size="small"
                variant="contained"
                startIcon={<SaveIcon />}
                onClick={handleSaveName}
              >
                Save
              </Button>
              <Button
                size="small"
                variant="outlined"
                startIcon={<CancelIcon />}
                onClick={handleCancelEditName}
              >
                Cancel
              </Button>
            </Box>
          </Box>
        ) : currentNameEntry ? (
          <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
            <Chip
              label={currentNameEntry.name}
              color="info"
              size="small"
              onDelete={handleRemoveName}
            />
            <Tooltip title="Edit name">
              <IconButton size="small" onClick={handleStartEditName}>
                <EditIcon fontSize="small" />
              </IconButton>
            </Tooltip>
            {currentNameEntry.cellKey !== node.id && (
              <Typography variant="caption" color="text.secondary">
                (from {currentNameEntry.cellKey})
              </Typography>
            )}
          </Box>
        ) : node.excelName ? (
          <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
            <Chip
              label={node.excelName}
              color="secondary"
              size="small"
              variant="outlined"
            />
            <Typography variant="caption" color="text.secondary">
              (from Excel)
            </Typography>
            <Tooltip title="Override with custom name">
              <IconButton size="small" onClick={handleStartEditName}>
                <EditIcon fontSize="small" />
              </IconButton>
            </Tooltip>
          </Box>
        ) : (
          <Button
            size="small"
            variant="outlined"
            startIcon={<AddIcon />}
            onClick={handleStartEditName}
          >
            Add Name
          </Button>
        )}
      </Box>

      {/* Cell Notes Section */}
      <Box sx={{ mb: 2 }}>
        <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 1 }}>
          <NotesIcon fontSize="small" color="warning" />
          <Typography variant="subtitle2" color="text.secondary">
            Notes ({cellNotes.length})
          </Typography>
        </Box>

        {cellNotes.map((entry, idx) => {
          const otherCells = getOtherCellsForNote(entry);
          const isEditing =
            editingNote?.cellKey === entry.cellKey &&
            editingNote?.oldNote === entry.note;

          return (
            <Paper
              key={`${entry.cellKey}-${idx}`}
              variant="outlined"
              sx={{ p: 1, mb: 1, bgcolor: "action.hover" }}
            >
              {isEditing ? (
                <Box>
                  <TextField
                    size="small"
                    fullWidth
                    multiline
                    rows={2}
                    value={editNoteValue}
                    onChange={(e) => setEditNoteValue(e.target.value)}
                    autoFocus
                    sx={{ mb: 1 }}
                  />
                  <Box sx={{ display: "flex", gap: 1 }}>
                    <Button
                      size="small"
                      variant="contained"
                      onClick={handleSaveEditNote}
                    >
                      Save
                    </Button>
                    <Button
                      size="small"
                      variant="outlined"
                      onClick={handleCancelEditNote}
                    >
                      Cancel
                    </Button>
                  </Box>
                </Box>
              ) : (
                <>
                  <Typography variant="body2" sx={{ mb: 0.5 }}>
                    {entry.note}
                  </Typography>
                  {otherCells.length > 0 && (
                    <Alert severity="info" sx={{ py: 0, mb: 0.5 }} icon={false}>
                      <Typography variant="caption">
                        Also applies to: {otherCells.slice(0, 3).join(", ")}
                        {otherCells.length > 3 &&
                          ` +${otherCells.length - 3} more`}
                      </Typography>
                    </Alert>
                  )}
                  <Box sx={{ display: "flex", gap: 0.5 }}>
                    <Tooltip title="Edit note">
                      <IconButton
                        size="small"
                        onClick={() => handleStartEditNote(entry)}
                      >
                        <EditIcon fontSize="small" />
                      </IconButton>
                    </Tooltip>
                    <Tooltip title="Delete note">
                      <IconButton
                        size="small"
                        color="error"
                        onClick={() => handleRemoveNote(entry)}
                      >
                        <DeleteIcon fontSize="small" />
                      </IconButton>
                    </Tooltip>
                  </Box>
                </>
              )}
            </Paper>
          );
        })}

        {isAddingNote ? (
          <Box>
            <TextField
              size="small"
              fullWidth
              multiline
              rows={2}
              value={newNoteValue}
              onChange={(e) => setNewNoteValue(e.target.value)}
              placeholder="Enter a note"
              autoFocus
              sx={{ mb: 1 }}
            />
            <Box sx={{ display: "flex", gap: 1 }}>
              <Button
                size="small"
                variant="contained"
                startIcon={<SaveIcon />}
                onClick={handleSaveNewNote}
              >
                Save
              </Button>
              <Button
                size="small"
                variant="outlined"
                startIcon={<CancelIcon />}
                onClick={handleCancelAddNote}
              >
                Cancel
              </Button>
            </Box>
          </Box>
        ) : (
          <Button
            size="small"
            variant="outlined"
            startIcon={<AddIcon />}
            onClick={handleStartAddNote}
          >
            Add Note
          </Button>
        )}
      </Box>

      {node.formula && (
        <Box sx={{ mb: 2 }}>
          <Typography variant="subtitle2" color="text.secondary">
            Formula
          </Typography>
          <Paper
            variant="outlined"
            sx={{
              p: 1,
              bgcolor: "grey.900",
              fontFamily: "monospace",
              fontSize: "0.875rem",
              wordBreak: "break-all",
            }}
          >
            {node.formula}
          </Paper>
        </Box>
      )}

      {node.value !== undefined && (
        <Box sx={{ mb: 2 }}>
          <Typography variant="subtitle2" color="text.secondary">
            Value
          </Typography>
          <Typography variant="body1">{formatValue(node.value)}</Typography>
        </Box>
      )}

      <Box sx={{ mb: 2 }}>
        <Typography variant="subtitle2" color="text.secondary">
          Type
        </Typography>
        <Chip
          label={node.hasFormula ? "Formula Cell" : "Referenced Cell"}
          size="small"
          color={node.hasFormula ? "success" : "default"}
        />
      </Box>

      <Divider sx={{ my: 2 }} />

      {/* Inputs Section */}
      <Accordion
        defaultExpanded={inputNodes.length > 0}
        disabled={inputNodes.length === 0}
        sx={{ bgcolor: "transparent", boxShadow: "none" }}
      >
        <AccordionSummary expandIcon={<ExpandMoreIcon />} sx={{ px: 0 }}>
          <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
            <InputIcon fontSize="small" color="info" />
            <Typography variant="subtitle2">
              Inputs ({inputNodes.length})
            </Typography>
          </Box>
        </AccordionSummary>
        <AccordionDetails sx={{ px: 0, pt: 0 }}>
          {inputNodes.length === 0 ? (
            <Typography variant="body2" color="text.secondary">
              No input references
            </Typography>
          ) : (
            <List dense disablePadding>
              {inputNodes.map((info) =>
                renderRelatedNode(info, expandedInputs.has(info.node.id), () =>
                  toggleInputExpanded(info.node.id)
                )
              )}
            </List>
          )}
        </AccordionDetails>
      </Accordion>

      {/* Outputs Section */}
      <Accordion
        defaultExpanded={outputNodes.length > 0}
        disabled={outputNodes.length === 0}
        sx={{ bgcolor: "transparent", boxShadow: "none" }}
      >
        <AccordionSummary expandIcon={<ExpandMoreIcon />} sx={{ px: 0 }}>
          <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
            <OutputIcon fontSize="small" color="warning" />
            <Typography variant="subtitle2">
              Outputs ({outputNodes.length})
            </Typography>
          </Box>
        </AccordionSummary>
        <AccordionDetails sx={{ px: 0, pt: 0 }}>
          {outputNodes.length === 0 ? (
            <Typography variant="body2" color="text.secondary">
              Not referenced by other cells
            </Typography>
          ) : (
            <List dense disablePadding>
              {outputNodes.map((info) =>
                renderRelatedNode(info, expandedOutputs.has(info.node.id), () =>
                  toggleOutputExpanded(info.node.id)
                )
              )}
            </List>
          )}
        </AccordionDetails>
      </Accordion>
    </Paper>
  );
}
