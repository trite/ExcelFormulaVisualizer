import type { GraphData, GraphNode, GraphEdge } from "../types";

/**
 * Options for filtering the graph
 */
export interface GraphFilterOptions {
  maxNodes: number;
  showOnlyFormulas: boolean;
  selectedSheets: string[] | "all";
  prioritizeNamed: boolean;
  namedCells: Set<string>;
}

/**
 * Statistics about the filtering result
 */
export interface GraphFilterStats {
  totalNodes: number;
  totalEdges: number;
  visibleNodes: number;
  visibleEdges: number;
  hiddenBySheet: number;
  hiddenByFormula: number;
  hiddenByLimit: number;
  limitReached: boolean;
}

/**
 * Result of filtering operation
 */
export interface FilteredGraphResult {
  graphData: GraphData;
  stats: GraphFilterStats;
}

/**
 * Node with computed metrics for importance scoring
 */
interface NodeWithMetrics extends GraphNode {
  inDegree: number;
  outDegree: number;
  isCrossSheet: boolean;
  importanceScore: number;
}

/**
 * Default filter options
 */
export const DEFAULT_FILTER_OPTIONS: GraphFilterOptions = {
  maxNodes: 1000,
  showOnlyFormulas: false,
  selectedSheets: "all",
  prioritizeNamed: true,
  namedCells: new Set(),
};

/**
 * Calculate in-degree and out-degree for each node
 */
function calculateNodeMetrics(
  nodes: GraphNode[],
  edges: GraphEdge[]
): Map<string, { inDegree: number; outDegree: number; crossSheetEdges: Set<string> }> {
  const metrics = new Map<string, { inDegree: number; outDegree: number; crossSheetEdges: Set<string> }>();

  // Initialize metrics for all nodes
  for (const node of nodes) {
    metrics.set(node.id, { inDegree: 0, outDegree: 0, crossSheetEdges: new Set() });
  }

  // Count degrees from edges
  for (const edge of edges) {
    const sourceMetrics = metrics.get(edge.source);
    const targetMetrics = metrics.get(edge.target);

    if (sourceMetrics) {
      sourceMetrics.outDegree++;
    }
    if (targetMetrics) {
      targetMetrics.inDegree++;
    }

    // Track cross-sheet edges
    const sourceSheet = edge.source.split("!")[0];
    const targetSheet = edge.target.split("!")[0];
    if (sourceSheet !== targetSheet) {
      if (sourceMetrics) {
        sourceMetrics.crossSheetEdges.add(edge.target);
      }
      if (targetMetrics) {
        targetMetrics.crossSheetEdges.add(edge.source);
      }
    }
  }

  return metrics;
}

/**
 * Calculate importance score for a node
 * Higher scores = more important nodes to show
 */
function calculateImportanceScore(
  node: GraphNode,
  metrics: { inDegree: number; outDegree: number; crossSheetEdges: Set<string> },
  isNamed: boolean
): number {
  let score = 0;

  // In-degree weighted more heavily (cells that many formulas depend on are important)
  score += metrics.inDegree * 2;

  // Out-degree (cells that reference many others)
  score += metrics.outDegree;

  // Cross-sheet references are often key interfaces
  if (metrics.crossSheetEdges.size > 0) {
    score += 5;
  }

  // Named cells are explicitly marked as important by user
  if (isNamed) {
    score += 10;
  }

  // Formula cells are computation hubs
  if (node.hasFormula) {
    score += 3;
  }

  return score;
}

/**
 * Filter graph based on options, applying importance-based node limiting
 */
export function filterGraph(
  graphData: GraphData,
  options: GraphFilterOptions
): FilteredGraphResult {
  const { maxNodes, showOnlyFormulas, selectedSheets, prioritizeNamed, namedCells } = options;

  const stats: GraphFilterStats = {
    totalNodes: graphData.nodes.length,
    totalEdges: graphData.edges.length,
    visibleNodes: 0,
    visibleEdges: 0,
    hiddenBySheet: 0,
    hiddenByFormula: 0,
    hiddenByLimit: 0,
    limitReached: false,
  };

  // Early return if graph is empty
  if (graphData.nodes.length === 0) {
    return {
      graphData: { nodes: [], edges: [] },
      stats,
    };
  }

  // Calculate metrics for all nodes
  const nodeMetrics = calculateNodeMetrics(graphData.nodes, graphData.edges);

  // Step 1: Apply sheet filter
  let filteredNodes = graphData.nodes;
  if (selectedSheets !== "all" && selectedSheets.length > 0) {
    const sheetSet = new Set(selectedSheets);
    const beforeCount = filteredNodes.length;
    filteredNodes = filteredNodes.filter((node) => sheetSet.has(node.sheet));
    stats.hiddenBySheet = beforeCount - filteredNodes.length;
  }

  // Step 2: Apply "show only formulas" filter
  // Keep nodes that either have formulas OR are referenced by other cells (inDegree > 0)
  if (showOnlyFormulas) {
    const beforeCount = filteredNodes.length;
    filteredNodes = filteredNodes.filter((node) => {
      const metrics = nodeMetrics.get(node.id);
      // Keep if it has a formula OR if other cells reference it
      return node.hasFormula || (metrics && metrics.inDegree > 0);
    });
    stats.hiddenByFormula = beforeCount - filteredNodes.length;
  }

  // Step 3: Calculate importance scores and sort
  const nodesWithScores: NodeWithMetrics[] = filteredNodes.map((node) => {
    const metrics = nodeMetrics.get(node.id) || {
      inDegree: 0,
      outDegree: 0,
      crossSheetEdges: new Set<string>(),
    };
    const isNamed = prioritizeNamed && namedCells.has(node.id);

    return {
      ...node,
      inDegree: metrics.inDegree,
      outDegree: metrics.outDegree,
      isCrossSheet: metrics.crossSheetEdges.size > 0,
      importanceScore: calculateImportanceScore(node, metrics, isNamed),
    };
  });

  // Sort by importance score descending
  nodesWithScores.sort((a, b) => b.importanceScore - a.importanceScore);

  // Step 4: Apply node limit
  let finalNodes: NodeWithMetrics[];
  if (nodesWithScores.length > maxNodes) {
    stats.limitReached = true;
    stats.hiddenByLimit = nodesWithScores.length - maxNodes;

    // If prioritizing named cells, ensure they're included
    if (prioritizeNamed && namedCells.size > 0) {
      const namedNodes: NodeWithMetrics[] = [];
      const otherNodes: NodeWithMetrics[] = [];

      for (const node of nodesWithScores) {
        if (namedCells.has(node.id)) {
          namedNodes.push(node);
        } else {
          otherNodes.push(node);
        }
      }

      // Take all named nodes (up to limit) + fill rest with highest scored
      const namedToTake = Math.min(namedNodes.length, maxNodes);
      const othersToTake = maxNodes - namedToTake;

      finalNodes = [
        ...namedNodes.slice(0, namedToTake),
        ...otherNodes.slice(0, othersToTake),
      ];
    } else {
      finalNodes = nodesWithScores.slice(0, maxNodes);
    }
  } else {
    finalNodes = nodesWithScores;
  }

  // Step 5: Filter edges to only those connecting visible nodes
  const visibleNodeIds = new Set(finalNodes.map((n) => n.id));
  const filteredEdges = graphData.edges.filter(
    (edge) => visibleNodeIds.has(edge.source) && visibleNodeIds.has(edge.target)
  );

  // Strip the metrics from the final nodes (keep original GraphNode shape)
  const cleanNodes: GraphNode[] = finalNodes.map((nodeWithMetrics) => {
    const { inDegree, outDegree, isCrossSheet, importanceScore, ...node } = nodeWithMetrics;
    // Suppress unused variable warnings - we're intentionally stripping these
    void inDegree; void outDegree; void isCrossSheet; void importanceScore;
    return node;
  });

  stats.visibleNodes = cleanNodes.length;
  stats.visibleEdges = filteredEdges.length;

  return {
    graphData: {
      nodes: cleanNodes,
      edges: filteredEdges,
    },
    stats,
  };
}

/**
 * Create a filter options object with defaults
 */
export function createFilterOptions(
  overrides: Partial<GraphFilterOptions> = {}
): GraphFilterOptions {
  return {
    ...DEFAULT_FILTER_OPTIONS,
    ...overrides,
    // Ensure namedCells is always a Set
    namedCells: overrides.namedCells || new Set(),
  };
}
