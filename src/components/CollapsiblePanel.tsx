import { useState, useEffect, useCallback } from "react";
import { Box, IconButton, Typography, Paper } from "@mui/material";
import ChevronLeftIcon from "@mui/icons-material/ChevronLeft";
import ChevronRightIcon from "@mui/icons-material/ChevronRight";

interface CollapsiblePanelProps {
  title: string;
  side: "left" | "right";
  storageKey: string;
  defaultExpanded?: boolean;
  width?: number;
  collapsedWidth?: number;
  children: React.ReactNode;
}

export function CollapsiblePanel({
  title,
  side,
  storageKey,
  defaultExpanded = true,
  width = 300,
  collapsedWidth = 40,
  children,
}: CollapsiblePanelProps) {
  // Load initial state from localStorage
  const [isExpanded, setIsExpanded] = useState(() => {
    const stored = localStorage.getItem(`panel-${storageKey}`);
    return stored !== null ? stored === "true" : defaultExpanded;
  });

  // Persist state to localStorage
  useEffect(() => {
    localStorage.setItem(`panel-${storageKey}`, String(isExpanded));
  }, [isExpanded, storageKey]);

  // Keyboard shortcut handler
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Skip if user is typing in an input
      if (
        e.target instanceof HTMLInputElement ||
        e.target instanceof HTMLTextAreaElement
      ) {
        return;
      }

      // '[' toggles left panel, ']' toggles right panel
      if (e.key === "[" && side === "left") {
        setIsExpanded((prev) => !prev);
      } else if (e.key === "]" && side === "right") {
        setIsExpanded((prev) => !prev);
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [side]);

  const handleToggle = useCallback(() => {
    setIsExpanded((prev) => !prev);
  }, []);

  const ExpandIcon = side === "left" ? ChevronRightIcon : ChevronLeftIcon;
  const CollapseIcon = side === "left" ? ChevronLeftIcon : ChevronRightIcon;

  return (
    <Paper
      sx={{
        width: isExpanded ? width : collapsedWidth,
        minWidth: isExpanded ? width : collapsedWidth,
        height: "100%",
        display: "flex",
        flexDirection: "column",
        overflow: "hidden",
        transition: "width 200ms ease-in-out, min-width 200ms ease-in-out",
        flexShrink: 0,
      }}
    >
      {isExpanded ? (
        // Expanded view
        <>
          <Box
            sx={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              p: 1,
              borderBottom: 1,
              borderColor: "divider",
              minHeight: 48,
            }}
          >
            <Typography variant="subtitle2" fontWeight={600}>
              {title}
            </Typography>
            <IconButton size="small" onClick={handleToggle} title={`Collapse (${side === "left" ? "[" : "]"})`}>
              <CollapseIcon fontSize="small" />
            </IconButton>
          </Box>
          <Box sx={{ flexGrow: 1, overflow: "auto" }}>{children}</Box>
        </>
      ) : (
        // Collapsed view - just show expand button
        <Box
          sx={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            pt: 1,
          }}
        >
          <IconButton size="small" onClick={handleToggle} title={`Expand ${title} (${side === "left" ? "[" : "]"})`}>
            <ExpandIcon fontSize="small" />
          </IconButton>
          {/* Rotated title for collapsed state */}
          <Typography
            variant="caption"
            sx={{
              writingMode: "vertical-rl",
              textOrientation: "mixed",
              transform: side === "left" ? "rotate(180deg)" : "none",
              mt: 2,
              color: "text.secondary",
              userSelect: "none",
            }}
          >
            {title}
          </Typography>
        </Box>
      )}
    </Paper>
  );
}
