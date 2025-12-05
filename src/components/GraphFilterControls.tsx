import {
  Box,
  Slider,
  FormControlLabel,
  Checkbox,
  Typography,
  Tooltip,
  Alert,
} from "@mui/material";
import WarningAmberIcon from "@mui/icons-material/WarningAmber";
import type { GraphFilterStats } from "../types";

interface GraphFilterControlsProps {
  maxNodes: number;
  showOnlyFormulas: boolean;
  stats: GraphFilterStats;
  onMaxNodesChange: (value: number) => void;
  onShowOnlyFormulasChange: (value: boolean) => void;
}

// Slider marks for common values
const sliderMarks = [
  { value: 100, label: "100" },
  { value: 500, label: "500" },
  { value: 1000, label: "1K" },
  { value: 2000, label: "2K" },
  { value: 5000, label: "5K" },
];

export function GraphFilterControls({
  maxNodes,
  showOnlyFormulas,
  stats,
  onMaxNodesChange,
  onShowOnlyFormulasChange,
}: GraphFilterControlsProps) {
  const percentVisible =
    stats.totalNodes > 0
      ? Math.round((stats.visibleNodes / stats.totalNodes) * 100)
      : 100;

  const showWarning = stats.limitReached && percentVisible < 50;

  return (
    <Box sx={{ display: "flex", alignItems: "center", gap: 2, flexWrap: "wrap" }}>
      {/* Max nodes slider */}
      <Box sx={{ display: "flex", alignItems: "center", gap: 1, minWidth: 200 }}>
        <Typography variant="body2" color="text.secondary" sx={{ whiteSpace: "nowrap" }}>
          Max nodes:
        </Typography>
        <Slider
          value={maxNodes}
          onChange={(_, value) => onMaxNodesChange(value as number)}
          min={100}
          max={5000}
          step={100}
          marks={sliderMarks}
          valueLabelDisplay="auto"
          size="small"
          sx={{ minWidth: 150 }}
        />
      </Box>

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
  );
}
