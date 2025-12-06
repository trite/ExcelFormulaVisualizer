import { useRef, useEffect, useState, useCallback, useMemo } from "react";
import ForceGraph2D from "react-force-graph-2d";
import {
  Box,
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
  selectedNodes: Set<string>;
  onSelectionChange: (nodes: Set<string>) => void;
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
  selectedNodes,
  onSelectionChange,
  onNodeSelect,
  cellNameMap,
}: ForceGraphViewProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const graphRef = useRef<ForceGraphInstance>(null);
  const [dimensions, setDimensions] = useState({ width: 800, height: 600 });

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

  const onSelectionChangeRef = useRef(onSelectionChange);
  onSelectionChangeRef.current = onSelectionChange;

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

  // Reset initial fit when data changes
  useEffect(() => {
    hasInitialFit.current = false;
  }, [data]);

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

  // Center on selected nodes when selection changes externally
  useEffect(() => {
    if (selectedNodes.size > 0 && graphRef.current) {
      // Center on the first selected node
      const firstSelectedId = Array.from(selectedNodes)[0];
      const node = data.nodes.find((n) => n.id === firstSelectedId);
      if (node && node.x !== undefined && node.y !== undefined) {
        graphRef.current.centerAt(node.x, node.y, 300);
      }
    }
  }, [selectedNodes, data.nodes]);

  // Emit particles periodically for pulse effect (every 2 seconds)
  useEffect(() => {
    const emitPulse = () => {
      if (document.hidden) return;

      if (graphRef.current && data.links.length > 0) {
        data.links.forEach((link) => {
          graphRef.current.emitParticle(link);
        });
      }
    };

    const initialTimeout = setTimeout(emitPulse, 500);
    const interval = setInterval(emitPulse, 2000);

    return () => {
      clearTimeout(initialTimeout);
      clearInterval(interval);
    };
  }, [data.links]);

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

      data.links.forEach((link) => {
        const sourceId = getLinkNodeId(link.source);
        const targetId = getLinkNodeId(link.target);
        const linkKey = `${sourceId}->${targetId}`;

        if (targetId === selectedId) {
          inputs.add(sourceId);
          inLinks.add(linkKey);
        }
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
    }, [selectedNodes, data.links]);

  const hasActiveSelection = selectedNodes.size === 1;

  const getNodeColor = useCallback(
    (node: ForceGraphNode) => {
      if (
        hasActiveSelection &&
        !selectedNodes.has(node.id) &&
        !inputNodeIds.has(node.id) &&
        !outputNodeIds.has(node.id)
      ) {
        const baseColor = sheetColorMap[node.sheet] || "#90caf9";
        return baseColor + "40";
      }
      return sheetColorMap[node.sheet] || "#90caf9";
    },
    [sheetColorMap, selectedNodes, inputNodeIds, outputNodeIds, hasActiveSelection]
  );

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

  const getLinkWidth = useCallback(
    (link: { source: unknown; target: unknown }) => {
      if (!hasActiveSelection) {
        return 1;
      }

      const sourceId = getLinkNodeId(link.source);
      const targetId = getLinkNodeId(link.target);
      const linkKey = `${sourceId}->${targetId}`;

      if (inputLinkKeys.has(linkKey) || outputLinkKeys.has(linkKey)) {
        return 2;
      }
      return 0.5;
    },
    [hasActiveSelection, inputLinkKeys, outputLinkKeys]
  );

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

  // Handle node click with multi-selection support
  const handleNodeClick = useCallback(
    (node: ForceGraphNode, event: MouseEvent) => {
      if (event.ctrlKey || event.shiftKey || event.altKey) {
        // Multi-selection mode: toggle node in selection
        const newSet = new Set(selectedNodes);
        if (newSet.has(node.id)) {
          newSet.delete(node.id);
        } else {
          newSet.add(node.id);
        }
        onSelectionChangeRef.current(newSet);
      } else {
        // Single selection mode
        const wasOnlySelected = selectedNodes.has(node.id) && selectedNodes.size === 1;
        if (wasOnlySelected) {
          onSelectionChangeRef.current(new Set());
          onNodeSelectRef.current(null);
        } else {
          onSelectionChangeRef.current(new Set([node.id]));
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

  const handleBackgroundClick = useCallback(() => {
    onSelectionChangeRef.current(new Set());
    onNodeSelectRef.current(null);
  }, []);

  const handleNodeDrag = useCallback(
    (node: ForceGraphNode, translate: { x: number; y: number }) => {
      if (selectedNodes.has(node.id)) {
        data.nodes
          .filter((n) => selectedNodes.has(n.id) && n.id !== node.id)
          .forEach((n) => {
            n.fx = (n.x || 0) + translate.x;
            n.fy = (n.y || 0) + translate.y;
          });
      }
    },
    [selectedNodes, data.nodes]
  );

  const handleNodeDragEnd = useCallback(
    (node: ForceGraphNode) => {
      if (selectedNodes.has(node.id)) {
        data.nodes
          .filter((n) => selectedNodes.has(n.id))
          .forEach((n) => {
            n.fx = undefined;
            n.fy = undefined;
          });
      }
    },
    [selectedNodes, data.nodes]
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

      setSelectionBox((prev: SelectionBox | null) => (prev ? { ...prev, endX: x, endY: y } : null));
    },
    [isDragging, selectionBox]
  );

  const handleOverlayMouseUp = useCallback(() => {
    if (!isDragging || !selectionBox || !graphRef.current) {
      setIsDragging(false);
      setSelectionBox(null);
      return;
    }

    const minX = Math.min(selectionBox.startX, selectionBox.endX);
    const maxX = Math.max(selectionBox.startX, selectionBox.endX);
    const minY = Math.min(selectionBox.startY, selectionBox.endY);
    const maxY = Math.max(selectionBox.startY, selectionBox.endY);

    const graph = graphRef.current;
    const nodesInBox: string[] = [];

    data.nodes.forEach((node) => {
      if (node.x !== undefined && node.y !== undefined) {
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

    if (nodesInBox.length > 0) {
      const newSet = new Set(selectedNodes);
      nodesInBox.forEach((id) => newSet.add(id));
      onSelectionChangeRef.current(newSet);
    }

    setIsDragging(false);
    setSelectionBox(null);
  }, [isDragging, selectionBox, data.nodes, selectedNodes]);

  // Draw custom node with label and glow effects
  const drawNode = useCallback(
    (
      node: ForceGraphNode,
      ctx: CanvasRenderingContext2D,
      globalScale: number
    ) => {
      const userDefinedName = cellNameMap?.get(node.id);
      const cellName = userDefinedName || node.excelName;
      const fontSize = 12 / globalScale;
      const nodeRadius = node.hasFormula ? 6 : 4;
      const color = getNodeColor(node);
      const isSelected = selectedNodes.has(node.id);
      const isInput = hasActiveSelection && inputNodeIds.has(node.id);
      const isOutput = hasActiveSelection && outputNodeIds.has(node.id);
      const isNeighbor = isInput || isOutput;
      const isDimmed = hasActiveSelection && !isSelected && !isNeighbor;

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

      ctx.save();

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

      ctx.beginPath();
      ctx.arc(node.x || 0, node.y || 0, nodeRadius, 0, 2 * Math.PI);
      ctx.fillStyle = color;
      ctx.fill();

      if (isSelected || isInput || isOutput) {
        ctx.shadowBlur = 25;
        ctx.fill();
      }

      ctx.restore();

      if (isSelected) {
        ctx.beginPath();
        ctx.arc(node.x || 0, node.y || 0, nodeRadius, 0, 2 * Math.PI);
        ctx.strokeStyle = "#ffffff";
        ctx.lineWidth = 2 / globalScale;
        ctx.stroke();
      }

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
      {/* Toolbar with sheet legend and selection count */}
      <Box sx={{ p: 1, display: "flex", alignItems: "center", gap: 2 }}>
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
            {selectedNodes.size} node{selectedNodes.size !== 1 ? "s" : ""} selected
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
          graphData={data}
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
