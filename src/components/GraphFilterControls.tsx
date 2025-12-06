import { useState } from "react";
import {
  Box,
  FormControlLabel,
  Checkbox,
  Typography,
  Tooltip,
  Alert,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
} from "@mui/material";
import type { SelectChangeEvent } from "@mui/material";
import WarningAmberIcon from "@mui/icons-material/WarningAmber";
import type { GraphFilterStats } from "../types";

interface GraphFilterControlsProps {
  maxNodes: number;
  showOnlyFormulas: boolean;
  stats: GraphFilterStats;
  onMaxNodesChange: (value: number) => void;
  onShowOnlyFormulasChange: (value: boolean) => void;
}

// Dropdown options for max nodes
const MAX_NODE_OPTIONS = [
  { value: 50, label: "50" },
  { value: 250, label: "250" },
  { value: 1000, label: "1,000" },
  { value: 5000, label: "5,000" },
  { value: Infinity, label: "All" },
];

export function GraphFilterControls({
  maxNodes,
  showOnlyFormulas,
  stats,
  onMaxNodesChange,
  onShowOnlyFormulasChange,
}: GraphFilterControlsProps) {
  const [showAllWarning, setShowAllWarning] = useState(false);
  const [pendingAllChange, setPendingAllChange] = useState(false);

  const percentVisible =
    stats.totalNodes > 0
      ? Math.round((stats.visibleNodes / stats.totalNodes) * 100)
      : 100;

  const showWarning = stats.limitReached && percentVisible < 50;

  const handleMaxNodesChange = (event: SelectChangeEvent<number>) => {
    const value = event.target.value as number;

    // If selecting "All" and there are many nodes, show warning
    if (value === Infinity && stats.totalNodes > 5000) {
      setPendingAllChange(true);
      setShowAllWarning(true);
    } else {
      onMaxNodesChange(value);
    }
  };

  const handleConfirmAll = () => {
    setShowAllWarning(false);
    if (pendingAllChange) {
      onMaxNodesChange(Infinity);
      setPendingAllChange(false);
    }
  };

  const handleCancelAll = () => {
    setShowAllWarning(false);
    setPendingAllChange(false);
  };

  // Get the current display value (handle Infinity)
  const displayValue = maxNodes === Infinity ? Infinity : maxNodes;

  return (
    <>
      <Box sx={{ display: "flex", alignItems: "center", gap: 2, flexWrap: "wrap" }}>
        {/* Max nodes dropdown */}
        <FormControl size="small" sx={{ minWidth: 120 }}>
          <InputLabel>Max Nodes</InputLabel>
          <Select
            value={displayValue}
            label="Max Nodes"
            onChange={handleMaxNodesChange}
          >
            {MAX_NODE_OPTIONS.map((option) => (
              <MenuItem key={option.label} value={option.value}>
                {option.label}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {/* Show only formulas checkbox */}
        <FormControlLabel
          control={
            <Checkbox
              checked={showOnlyFormulas}
              onChange={(e) => onShowOnlyFormulasChange(e.target.checked)}
              size="small"
            />
          }
          label={
            <Typography variant="body2" color="text.secondary">
              Formulas only
            </Typography>
          }
          sx={{ mr: 0 }}
        />

        {/* Stats display */}
        <Tooltip
          title={
            stats.limitReached
              ? `Showing top ${stats.visibleNodes} nodes by importance. ${stats.hiddenByLimit} nodes hidden by limit.`
              : "All nodes are visible"
          }
        >
          <Typography
            variant="body2"
            sx={{
              color: stats.limitReached ? "warning.main" : "text.secondary",
              display: "flex",
              alignItems: "center",
              gap: 0.5,
            }}
          >
            {stats.limitReached && <WarningAmberIcon fontSize="small" />}
            {stats.visibleNodes.toLocaleString()} of {stats.totalNodes.toLocaleString()} nodes
            ({percentVisible}%)
          </Typography>
        </Tooltip>

        {/* Prominent warning when less than 50% visible */}
        {showWarning && (
          <Alert
            severity="warning"
            variant="outlined"
            sx={{
              py: 0,
              px: 1,
              "& .MuiAlert-message": { py: 0.5 },
            }}
          >
            Large file - increase limit or filter by sheet to see more
          </Alert>
        )}
      </Box>

      {/* Warning dialog for selecting "All" with large datasets */}
      <Dialog open={showAllWarning} onClose={handleCancelAll}>
        <DialogTitle>Show All Nodes?</DialogTitle>
        <DialogContent>
          <Alert severity="warning" sx={{ mb: 2 }}>
            This file has {stats.totalNodes.toLocaleString()} nodes. Displaying all of them
            may cause performance issues or freeze your browser.
          </Alert>
          <Typography>
            Are you sure you want to display all nodes?
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={handleCancelAll}>Cancel</Button>
          <Button onClick={handleConfirmAll} color="warning" variant="contained">
            Show All
          </Button>
        </DialogActions>
      </Dialog>
    </>
  );
}
