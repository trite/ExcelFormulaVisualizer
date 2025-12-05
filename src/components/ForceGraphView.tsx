import { useRef, useEffect, useState, useCallback, useMemo } from "react";
import ForceGraph2D from "react-force-graph-2d";
import {
  Box,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  Chip,
  Stack,
  Typography,
} from "@mui/material";
import type { ForceGraphData, ForceGraphNode, GraphNode } from "../types";
import {
  SHEET_COLORS,
  SELECTION_COLORS,
  LABEL_COLORS,
  LINK_COLORS,
} from "../theme/colors";

// Using any because ForceGraphMethods type has export issues
type ForceGraphInstance = any;

interface ForceGraphViewProps {
  data: ForceGraphData;
  sheets: string[];
  selectedNodeId?: string | null;
  onNodeSelect: (node: GraphNode | null) => void;
  cellNameMap?: Map<string, string>;
}

interface SelectionBox {
  startX: number;
  startY: number;
  endX: number;
  endY: number;
}

export function ForceGraphView({
  data,
  sheets,
  selectedNodeId,
  onNodeSelect,
  cellNameMap,
}: ForceGraphViewProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const graphRef = useRef<ForceGraphInstance>(null);
  const [dimensions, setDimensions] = useState({ width: 800, height: 600 });
  const [selectedSheet, setSelectedSheet] = useState<string>("all");
  const [selectedNodes, setSelectedNodes] = useState<Set<string>>(new Set());

  // Box selection state
  const [isShiftDown, setIsShiftDown] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [selectionBox, setSelectionBox] = useState<SelectionBox | null>(null);
  const overlayRef = useRef<HTMLDivElement>(null);

  // Track if initial fit has been done
  const hasInitialFit = useRef(false);

  // Store refs to avoid recreating callbacks
  const onNodeSelectRef = useRef(onNodeSelect);
  onNodeSelectRef.current = onNodeSelect;

  // Create stable color map
  const sheetColorMap = useMemo(() => {
    const map: Record<string, string> = {};
    sheets.forEach((sheet, index) => {
      map[sheet] = SHEET_COLORS[index % SHEET_COLORS.length];
    });
    return map;
  }, [sheets]);

  // Helper to get link source/target ID (force-graph mutates these to objects at runtime)
  const getLinkNodeId = (node: unknown): string => {
    if (typeof node === "string") return node;
    if (typeof node === "object" && node !== null && "id" in node) {
      return (node as { id: string }).id;
    }
    return String(node);
  };

  // Filter data based on selected sheet
  const filteredData = useMemo(() => {
    // Reset initial fit when data changes
    hasInitialFit.current = false;

    if (selectedSheet === "all") {
      return data;
    }
    const filteredNodes = data.nodes.filter((n) => n.sheet === selectedSheet);
    const nodeIds = new Set(filteredNodes.map((n) => n.id));
    const filteredLinks = data.links.filter(
      (l) =>
        nodeIds.has(getLinkNodeId(l.source)) &&
        nodeIds.has(getLinkNodeId(l.target))
    );
    return { nodes: filteredNodes, links: filteredLinks };
  }, [data, selectedSheet]);

  // Handle container resize
  useEffect(() => {
    const updateDimensions = () => {
      if (containerRef.current) {
        const rect = containerRef.current.getBoundingClientRect();
        setDimensions({ width: rect.width, height: rect.height - 60 });
      }
    };

    updateDimensions();
    const observer = new ResizeObserver(updateDimensions);
    if (containerRef.current) {
      observer.observe(containerRef.current);
    }

    return () => observer.disconnect();
  }, []);

  // Handle shift key state
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Shift") {
        setIsShiftDown(true);
      }
    };
    const handleKeyUp = (e: KeyboardEvent) => {
      if (e.key === "Shift") {
        setIsShiftDown(false);
        // Cancel any ongoing selection if shift is released
        if (isDragging) {
          setIsDragging(false);
          setSelectionBox(null);
        }
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    window.addEventListener("keyup", handleKeyUp);
    return () => {
      window.removeEventListener("keydown", handleKeyDown);
      window.removeEventListener("keyup", handleKeyUp);
    };
  }, [isDragging]);

  // Sync selection when selectedNodeId changes externally
  useEffect(() => {
    if (selectedNodeId) {
      setSelectedNodes(new Set([selectedNodeId]));
      // Center on the selected node if it exists
      if (graphRef.current) {
        const node = filteredData.nodes.find((n) => n.id === selectedNodeId);
        if (node && node.x !== undefined && node.y !== undefined) {
          graphRef.current.centerAt(node.x, node.y, 300);
        }
      }
    } else {
      setSelectedNodes(new Set());
    }
  }, [selectedNodeId, filteredData.nodes]);

  // Emit particles periodically for pulse effect (every 2 seconds)
  useEffect(() => {
    const emitPulse = () => {
      if (graphRef.current && filteredData.links.length > 0) {
        // Emit a particle on each link
        filteredData.links.forEach((link) => {
          graphRef.current.emitParticle(link);
        });
      }
    };

    // Emit initial pulse after a short delay
    const initialTimeout = setTimeout(emitPulse, 500);
    // Then emit every 2 seconds
    const interval = setInterval(emitPulse, 2000);

    return () => {
      clearTimeout(initialTimeout);
      clearInterval(interval);
    };
  }, [filteredData.links]);

  // Compute input/output neighbor nodes and connected links for the selected node
  const { inputNodeIds, outputNodeIds, inputLinkKeys, outputLinkKeys } =
    useMemo(() => {
      if (selectedNodes.size !== 1) {
        return {
          inputNodeIds: new Set<string>(),
          outputNodeIds: new Set<string>(),
          inputLinkKeys: new Set<string>(),
          outputLinkKeys: new Set<string>(),
        };
      }

      const selectedId = Array.from(selectedNodes)[0];
      const inputs = new Set<string>();
      const outputs = new Set<string>();
      const inLinks = new Set<string>();
      const outLinks = new Set<string>();

      filteredData.links.forEach((link) => {
        const sourceId = getLinkNodeId(link.source);
        const targetId = getLinkNodeId(link.target);
        const linkKey = `${sourceId}->${targetId}`;

        // If selected node is the target, source is an input
        if (targetId === selectedId) {
          inputs.add(sourceId);
          inLinks.add(linkKey);
        }
        // If selected node is the source, target is an output
        if (sourceId === selectedId) {
          outputs.add(targetId);
          outLinks.add(linkKey);
        }
      });

      return {
        inputNodeIds: inputs,
        outputNodeIds: outputs,
        inputLinkKeys: inLinks,
        outputLinkKeys: outLinks,
      };
    }, [selectedNodes, filteredData.links]);

  // Check if we have an active selection (for dimming logic)
  const hasActiveSelection = selectedNodes.size === 1;

  // Get node color based on sheet and neighbor state (selected nodes keep their sheet color)
  const getNodeColor = useCallback(
    (node: ForceGraphNode) => {
      // Dim non-connected nodes when there's an active selection
      if (
        hasActiveSelection &&
        !selectedNodes.has(node.id) &&
        !inputNodeIds.has(node.id) &&
        !outputNodeIds.has(node.id)
      ) {
        const baseColor = sheetColorMap[node.sheet] || "#90caf9";
        return baseColor + "40"; // Add alpha for dimming
      }
      return sheetColorMap[node.sheet] || "#90caf9";
    },
    [
      sheetColorMap,
      selectedNodes,
      inputNodeIds,
      outputNodeIds,
      hasActiveSelection,
    ]
  );

  // Get link color based on connection to selected node
  const getLinkColor = useCallback(
    (link: { source: unknown; target: unknown }) => {
      if (!hasActiveSelection) {
        return LINK_COLORS.default;
      }

      const sourceId = getLinkNodeId(link.source);
      const targetId = getLinkNodeId(link.target);
      const linkKey = `${sourceId}->${targetId}`;

      if (inputLinkKeys.has(linkKey)) {
        return SELECTION_COLORS.input;
      }
      if (outputLinkKeys.has(linkKey)) {
        return SELECTION_COLORS.output;
      }
      return LINK_COLORS.dimmed;
    },
    [hasActiveSelection, inputLinkKeys, outputLinkKeys]
  );

  // Get link width based on connection to selected node
  const getLinkWidth = useCallback(
    (link: { source: unknown; target: unknown }) => {
      if (!hasActiveSelection) {
        return 1;
      }

      const sourceId = getLinkNodeId(link.source);
      const targetId = getLinkNodeId(link.target);
      const linkKey = `${sourceId}->${targetId}`;

      if (inputLinkKeys.has(linkKey) || outputLinkKeys.has(linkKey)) {
        return 2; // Thicker for connected links
      }
      return 0.5; // Thinner for non-connected links
    },
    [hasActiveSelection, inputLinkKeys, outputLinkKeys]
  );

  // Get particle color matching link color
  const getParticleColor = useCallback(
    (link: { source: unknown; target: unknown }) => {
      if (!hasActiveSelection) {
        return "rgba(255,255,255,0.6)";
      }

      const sourceId = getLinkNodeId(link.source);
      const targetId = getLinkNodeId(link.target);
      const linkKey = `${sourceId}->${targetId}`;

      if (inputLinkKeys.has(linkKey)) {
        return SELECTION_COLORS.input;
      }
      if (outputLinkKeys.has(linkKey)) {
        return SELECTION_COLORS.output;
      }
      return LINK_COLORS.default;
    },
    [hasActiveSelection, inputLinkKeys, outputLinkKeys]
  );

  // Handle node click with multi-selection support (idiomatic approach from force-graph example)
  const handleNodeClick = useCallback(
    (node: ForceGraphNode, event: MouseEvent) => {
      if (event.ctrlKey || event.shiftKey || event.altKey) {
        // Multi-selection mode: toggle node in selection
        setSelectedNodes((prev) => {
          const newSet = new Set(prev);
          if (newSet.has(node.id)) {
            newSet.delete(node.id);
          } else {
            newSet.add(node.id);
          }
          return newSet;
        });
      } else {
        // Single selection mode
        const wasOnlySelected =
          selectedNodes.has(node.id) && selectedNodes.size === 1;
        if (wasOnlySelected) {
          // Clicking the only selected node deselects it
          setSelectedNodes(new Set());
          onNodeSelectRef.current(null);
        } else {
          // Select only this node
          setSelectedNodes(new Set([node.id]));
          onNodeSelectRef.current({
            id: node.id,
            label: node.label,
            sheet: node.sheet,
            address: node.address,
            formula: node.formula,
            value: node.value,
            hasFormula: node.hasFormula,
          });
        }
      }
    },
    [selectedNodes]
  );

  // Handle background click to clear selection
  const handleBackgroundClick = useCallback(() => {
    setSelectedNodes(new Set());
    onNodeSelectRef.current(null);
  }, []);

  // Handle multi-node drag (from force-graph example)
  const handleNodeDrag = useCallback(
    (node: ForceGraphNode, translate: { x: number; y: number }) => {
      if (selectedNodes.has(node.id)) {
        // Move all selected nodes together
        filteredData.nodes
          .filter((n) => selectedNodes.has(n.id) && n.id !== node.id)
          .forEach((n) => {
            n.fx = (n.x || 0) + translate.x;
            n.fy = (n.y || 0) + translate.y;
          });
      }
    },
    [selectedNodes, filteredData.nodes]
  );

  const handleNodeDragEnd = useCallback(
    (node: ForceGraphNode) => {
      if (selectedNodes.has(node.id)) {
        // Release fixed positions for all selected nodes
        filteredData.nodes
          .filter((n) => selectedNodes.has(n.id))
          .forEach((n) => {
            n.fx = undefined;
            n.fy = undefined;
          });
      }
    },
    [selectedNodes, filteredData.nodes]
  );

  // Box selection handlers
  const handleOverlayMouseDown = useCallback(
    (e: React.MouseEvent) => {
      if (!isShiftDown || !overlayRef.current) return;

      const rect = overlayRef.current.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;

      setIsDragging(true);
      setSelectionBox({ startX: x, startY: y, endX: x, endY: y });
    },
    [isShiftDown]
  );

  const handleOverlayMouseMove = useCallback(
    (e: React.MouseEvent) => {
      if (!isDragging || !selectionBox || !overlayRef.current) return;

      const rect = overlayRef.current.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;

      setSelectionBox((prev) => (prev ? { ...prev, endX: x, endY: y } : null));
    },
    [isDragging, selectionBox]
  );

  const handleOverlayMouseUp = useCallback(() => {
    if (!isDragging || !selectionBox || !graphRef.current) {
      setIsDragging(false);
      setSelectionBox(null);
      return;
    }

    // Calculate box bounds in screen coordinates
    const minX = Math.min(selectionBox.startX, selectionBox.endX);
    const maxX = Math.max(selectionBox.startX, selectionBox.endX);
    const minY = Math.min(selectionBox.startY, selectionBox.endY);
    const maxY = Math.max(selectionBox.startY, selectionBox.endY);

    // Find nodes within the selection box
    const graph = graphRef.current;
    const nodesInBox: string[] = [];

    filteredData.nodes.forEach((node) => {
      if (node.x !== undefined && node.y !== undefined) {
        // Convert graph coordinates to screen coordinates
        const screenCoords = graph.graph2ScreenCoords(node.x, node.y);
        if (
          screenCoords.x >= minX &&
          screenCoords.x <= maxX &&
          screenCoords.y >= minY &&
          screenCoords.y <= maxY
        ) {
          nodesInBox.push(node.id);
        }
      }
    });

    // Update selection
    if (nodesInBox.length > 0) {
      setSelectedNodes((prev) => {
        const newSet = new Set(prev);
        nodesInBox.forEach((id) => newSet.add(id));
        return newSet;
      });
    }

    setIsDragging(false);
    setSelectionBox(null);
  }, [isDragging, selectionBox, filteredData.nodes]);

  // Draw custom node with label and glow effects for input/output neighbors
  const drawNode = useCallback(
    (
      node: ForceGraphNode,
      ctx: CanvasRenderingContext2D,
      globalScale: number
    ) => {
      const cellName = cellNameMap?.get(node.id);
      const fontSize = 12 / globalScale;
      const nodeRadius = node.hasFormula ? 6 : 4;
      const color = getNodeColor(node);
      const isSelected = selectedNodes.has(node.id);
      const isInput = hasActiveSelection && inputNodeIds.has(node.id);
      const isOutput = hasActiveSelection && outputNodeIds.has(node.id);
      const isNeighbor = isInput || isOutput;
      const isDimmed = hasActiveSelection && !isSelected && !isNeighbor;

      // Determine label based on zoom level
      // When zoomed out (globalScale < 0.5), show only name for named cells, address for others
      // When zoomed in, show "name (address)" for named cells, just address for others
      let label: string;
      if (cellName) {
        if (globalScale < 0.5) {
          label = cellName;
        } else {
          label = `${cellName} (${node.address})`;
        }
      } else {
        label = node.address;
      }

      // Save context state before applying glow effects
      ctx.save();

      // Apply glow effect for selected, input, or output nodes
      // Use fixed pixel values (not scaled) so glow looks consistent at all zoom levels
      if (isSelected) {
        ctx.shadowColor = SELECTION_COLORS.selected;
        ctx.shadowBlur = 15;
      } else if (isInput) {
        ctx.shadowColor = SELECTION_COLORS.input;
        ctx.shadowBlur = 15;
      } else if (isOutput) {
        ctx.shadowColor = SELECTION_COLORS.output;
        ctx.shadowBlur = 15;
      }

      // Draw node circle (with glow if selected or neighbor)
      ctx.beginPath();
      ctx.arc(node.x || 0, node.y || 0, nodeRadius, 0, 2 * Math.PI);
      ctx.fillStyle = color;
      ctx.fill();

      // For selected/input/output nodes, draw a second pass with stronger glow for visibility
      if (isSelected || isInput || isOutput) {
        ctx.shadowBlur = 25;
        ctx.fill();
      }

      // Restore context to clear shadow effects before drawing other elements
      ctx.restore();

      // Draw selection ring if selected (solid white)
      if (isSelected) {
        ctx.beginPath();
        ctx.arc(node.x || 0, node.y || 0, nodeRadius, 0, 2 * Math.PI);
        ctx.strokeStyle = "#ffffff";
        ctx.lineWidth = 2 / globalScale;
        ctx.stroke();
      }

      // Draw label (dimmed if not selected/neighbor)
      // For named cells, use a slightly different color to make them stand out
      ctx.font = `${fontSize}px Sans-Serif`;
      ctx.textAlign = "center";
      ctx.textBaseline = "top";
      if (isDimmed) {
        ctx.fillStyle = LABEL_COLORS.dimmed;
      } else if (cellName) {
        ctx.fillStyle = LABEL_COLORS.named;
      } else {
        ctx.fillStyle = LABEL_COLORS.default;
      }
      ctx.fillText(label, node.x || 0, (node.y || 0) + nodeRadius + 2);
    },
    [
      getNodeColor,
      selectedNodes,
      hasActiveSelection,
      inputNodeIds,
      outputNodeIds,
      cellNameMap,
    ]
  );

  // Calculate selection box style
  const selectionBoxStyle = useMemo(() => {
    if (!selectionBox) return null;

    const left = Math.min(selectionBox.startX, selectionBox.endX);
    const top = Math.min(selectionBox.startY, selectionBox.endY);
    const width = Math.abs(selectionBox.endX - selectionBox.startX);
    const height = Math.abs(selectionBox.endY - selectionBox.startY);

    return {
      position: "absolute" as const,
      left,
      top,
      width,
      height,
      border: "2px dashed #90caf9",
      backgroundColor: "rgba(144, 202, 249, 0.2)",
      pointerEvents: "none" as const,
    };
  }, [selectionBox]);

  return (
    <Box
      ref={containerRef}
      sx={{ width: "100%", height: "100%", position: "relative" }}
    >
      <Box sx={{ p: 1, display: "flex", alignItems: "center", gap: 2 }}>
        <FormControl size="small" sx={{ minWidth: 150 }}>
          <InputLabel>Sheet Filter</InputLabel>
          <Select
            value={selectedSheet}
            label="Sheet Filter"
            onChange={(e) => setSelectedSheet(e.target.value)}
          >
            <MenuItem value="all">All Sheets</MenuItem>
            {sheets.map((sheet) => (
              <MenuItem key={sheet} value={sheet}>
                {sheet}
              </MenuItem>
            ))}
          </Select>
        </FormControl>
        <Stack direction="row" spacing={1}>
          {sheets.map((sheet) => (
            <Chip
              key={sheet}
              label={sheet}
              size="small"
              sx={{
                backgroundColor: sheetColorMap[sheet],
                color: "#000",
              }}
            />
          ))}
        </Stack>
        {selectedNodes.size > 0 && (
          <Typography
            variant="body2"
            sx={{ ml: "auto", color: "text.secondary" }}
          >
            {selectedNodes.size} node{selectedNodes.size !== 1 ? "s" : ""}{" "}
            selected
          </Typography>
        )}
      </Box>
      <Box
        sx={{
          position: "relative",
          width: dimensions.width,
          height: dimensions.height,
        }}
      >
        <ForceGraph2D
          ref={graphRef}
          graphData={filteredData}
          width={dimensions.width}
          height={dimensions.height}
          nodeId="id"
          nodeCanvasObject={drawNode}
          nodePointerAreaPaint={(node, color, ctx) => {
            const nodeRadius = (node as ForceGraphNode).hasFormula ? 6 : 4;
            ctx.beginPath();
            ctx.arc(node.x || 0, node.y || 0, nodeRadius + 4, 0, 2 * Math.PI);
            ctx.fillStyle = color;
            ctx.fill();
          }}
          linkDirectionalArrowLength={4}
          linkDirectionalArrowRelPos={0.95}
          linkDirectionalParticles={0}
          linkDirectionalParticleWidth={4}
          linkDirectionalParticleSpeed={0.012}
          linkDirectionalParticleColor={getParticleColor}
          linkColor={getLinkColor}
          linkWidth={getLinkWidth}
          onNodeClick={handleNodeClick}
          onBackgroundClick={handleBackgroundClick}
          onNodeDrag={handleNodeDrag}
          onNodeDragEnd={handleNodeDragEnd}
          enableNodeDrag={!isShiftDown}
          enablePanInteraction={!isShiftDown}
          cooldownTime={500}
          onEngineStop={() => {
            // Fit all nodes in view only on initial render
            if (graphRef.current && !hasInitialFit.current) {
              hasInitialFit.current = true;
              graphRef.current.zoomToFit(400, 50);
            }
          }}
        />
        {/* Overlay for box selection when shift is held */}
        <Box
          ref={overlayRef}
          sx={{
            position: "absolute",
            top: 0,
            left: 0,
            width: "100%",
            height: "100%",
            cursor: isShiftDown ? "crosshair" : "default",
            pointerEvents: isShiftDown ? "auto" : "none",
          }}
          onMouseDown={handleOverlayMouseDown}
          onMouseMove={handleOverlayMouseMove}
          onMouseUp={handleOverlayMouseUp}
          onMouseLeave={handleOverlayMouseUp}
        >
          {selectionBox && selectionBoxStyle && <Box sx={selectionBoxStyle} />}
        </Box>
      </Box>
    </Box>
  );
}
