import { useState, useMemo, useCallback, useEffect } from "react";
import {
  Box,
  TextField,
  FormGroup,
  FormControlLabel,
  Checkbox,
  RadioGroup,
  Radio,
  Typography,
  InputAdornment,
  IconButton,
  Collapse,
  Divider,
} from "@mui/material";
import SearchIcon from "@mui/icons-material/Search";
import ClearIcon from "@mui/icons-material/Clear";
import TuneIcon from "@mui/icons-material/Tune";
import type { GraphData, CellMetadata } from "../types";

interface CellSearchProps {
  graphData: GraphData | null;
  cellMetadata: CellMetadata;
  cellNameMap?: Map<string, string>;
  onSearchResults: (matchingIds: Set<string> | null) => void;
}

type SearchMode = "contains" | "fuzzy" | "regex";

interface SearchFields {
  value: boolean;
  formula: boolean;
  names: boolean;
  notes: boolean;
}

// Simple fuzzy match - checks if all characters of query appear in order in the text
function fuzzyMatch(text: string, query: string): boolean {
  const textLower = text.toLowerCase();
  const queryLower = query.toLowerCase();

  let textIndex = 0;
  for (let i = 0; i < queryLower.length; i++) {
    const found = textLower.indexOf(queryLower[i], textIndex);
    if (found === -1) return false;
    textIndex = found + 1;
  }
  return true;
}

// Debounce hook
function useDebounce<T>(value: T, delay: number): T {
  const [debouncedValue, setDebouncedValue] = useState(value);

  useEffect(() => {
    const handler = setTimeout(() => {
      setDebouncedValue(value);
    }, delay);

    return () => {
      clearTimeout(handler);
    };
  }, [value, delay]);

  return debouncedValue;
}

export function CellSearch({
  graphData,
  cellMetadata,
  cellNameMap,
  onSearchResults,
}: CellSearchProps) {
  const [searchQuery, setSearchQuery] = useState("");
  const [searchMode, setSearchMode] = useState<SearchMode>("contains");
  const [searchFields, setSearchFields] = useState<SearchFields>({
    value: true,
    formula: true,
    names: true,
    notes: true,
  });
  const [showOptions, setShowOptions] = useState(false);
  const [regexError, setRegexError] = useState<string | null>(null);

  const debouncedQuery = useDebounce(searchQuery, 300);

  // Build a map of cell IDs to their notes
  const cellNotesMap = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const entry of cellMetadata.notes) {
      // Handle comma-separated cell keys
      const cellKeys = entry.cellKey.split(",").map((k) => k.trim());
      for (const key of cellKeys) {
        if (!map.has(key)) {
          map.set(key, []);
        }
        map.get(key)!.push(entry.note);
      }
    }
    return map;
  }, [cellMetadata.notes]);

  // Perform search
  const matchingIds = useMemo(() => {
    if (!graphData || !debouncedQuery.trim()) {
      return null;
    }

    const query = debouncedQuery.trim();
    const matches = new Set<string>();

    // Prepare regex if in regex mode
    let regex: RegExp | null = null;
    if (searchMode === "regex") {
      try {
        regex = new RegExp(query, "i");
        setRegexError(null);
      } catch (e) {
        setRegexError((e as Error).message);
        return null;
      }
    }

    // Helper to check if a string matches the query
    const matchesQuery = (text: string | undefined | null): boolean => {
      if (!text) return false;
      const textStr = String(text);

      switch (searchMode) {
        case "contains":
          return textStr.toLowerCase().includes(query.toLowerCase());
        case "fuzzy":
          return fuzzyMatch(textStr, query);
        case "regex":
          return regex ? regex.test(textStr) : false;
        default:
          return false;
      }
    };

    for (const node of graphData.nodes) {
      let isMatch = false;

      // Check value
      if (searchFields.value && node.value !== undefined) {
        if (matchesQuery(String(node.value))) {
          isMatch = true;
        }
      }

      // Check formula
      if (!isMatch && searchFields.formula && node.formula) {
        if (matchesQuery(node.formula)) {
          isMatch = true;
        }
      }

      // Check names (both user-defined and Excel names)
      if (!isMatch && searchFields.names) {
        const userName = cellNameMap?.get(node.id);
        const excelName = node.excelName;

        if (userName && matchesQuery(userName)) {
          isMatch = true;
        } else if (excelName && matchesQuery(excelName)) {
          isMatch = true;
        }
      }

      // Check notes
      if (!isMatch && searchFields.notes) {
        const notes = cellNotesMap.get(node.id);
        if (notes) {
          for (const note of notes) {
            if (matchesQuery(note)) {
              isMatch = true;
              break;
            }
          }
        }
      }

      // Also check the cell address itself
      if (!isMatch && matchesQuery(node.id)) {
        isMatch = true;
      }
      if (!isMatch && matchesQuery(node.address)) {
        isMatch = true;
      }

      if (isMatch) {
        matches.add(node.id);
      }
    }

    return matches;
  }, [
    graphData,
    debouncedQuery,
    searchMode,
    searchFields,
    cellNameMap,
    cellNotesMap,
  ]);

  // Notify parent of search results
  useEffect(() => {
    onSearchResults(matchingIds);
  }, [matchingIds, onSearchResults]);

  const handleClear = useCallback(() => {
    setSearchQuery("");
    setRegexError(null);
  }, []);

  const handleFieldToggle = useCallback((field: keyof SearchFields) => {
    setSearchFields((prev) => ({ ...prev, [field]: !prev[field] }));
  }, []);

  const matchCount = matchingIds?.size ?? null;
  const hasQuery = searchQuery.trim().length > 0;

  return (
    <Box sx={{ p: 1 }}>
      {/* Search input */}
      <TextField
        fullWidth
        size="small"
        placeholder="Search cells..."
        value={searchQuery}
        onChange={(e) => setSearchQuery(e.target.value)}
        error={!!regexError}
        helperText={regexError}
        InputProps={{
          startAdornment: (
            <InputAdornment position="start">
              <SearchIcon fontSize="small" color="action" />
            </InputAdornment>
          ),
          endAdornment: (
            <InputAdornment position="end">
              {hasQuery && (
                <IconButton size="small" onClick={handleClear}>
                  <ClearIcon fontSize="small" />
                </IconButton>
              )}
              <IconButton
                size="small"
                onClick={() => setShowOptions((prev) => !prev)}
                color={showOptions ? "primary" : "default"}
              >
                <TuneIcon fontSize="small" />
              </IconButton>
            </InputAdornment>
          ),
        }}
      />

      {/* Match count */}
      {hasQuery && matchCount !== null && (
        <Typography
          variant="caption"
          color={matchCount > 0 ? "text.secondary" : "warning.main"}
          sx={{ display: "block", mt: 0.5, ml: 1 }}
        >
          {matchCount} {matchCount === 1 ? "match" : "matches"}
        </Typography>
      )}

      {/* Search options */}
      <Collapse in={showOptions}>
        <Box sx={{ mt: 1.5 }}>
          <Divider sx={{ mb: 1 }} />

          {/* Search fields */}
          <Typography variant="caption" color="text.secondary" sx={{ mb: 0.5, display: "block" }}>
            Search in:
          </Typography>
          <FormGroup row sx={{ mb: 1 }}>
            <FormControlLabel
              control={
                <Checkbox
                  checked={searchFields.value}
                  onChange={() => handleFieldToggle("value")}
                  size="small"
                />
              }
              label={<Typography variant="caption">Value</Typography>}
              sx={{ mr: 1 }}
            />
            <FormControlLabel
              control={
                <Checkbox
                  checked={searchFields.formula}
                  onChange={() => handleFieldToggle("formula")}
                  size="small"
                />
              }
              label={<Typography variant="caption">Formula</Typography>}
              sx={{ mr: 1 }}
            />
            <FormControlLabel
              control={
                <Checkbox
                  checked={searchFields.names}
                  onChange={() => handleFieldToggle("names")}
                  size="small"
                />
              }
              label={<Typography variant="caption">Names</Typography>}
              sx={{ mr: 1 }}
            />
            <FormControlLabel
              control={
                <Checkbox
                  checked={searchFields.notes}
                  onChange={() => handleFieldToggle("notes")}
                  size="small"
                />
              }
              label={<Typography variant="caption">Notes</Typography>}
            />
          </FormGroup>

          {/* Search mode */}
          <Typography variant="caption" color="text.secondary" sx={{ mb: 0.5, display: "block" }}>
            Mode:
          </Typography>
          <RadioGroup
            row
            value={searchMode}
            onChange={(e) => {
              setSearchMode(e.target.value as SearchMode);
              setRegexError(null);
            }}
          >
            <FormControlLabel
              value="contains"
              control={<Radio size="small" />}
              label={<Typography variant="caption">Contains</Typography>}
            />
            <FormControlLabel
              value="fuzzy"
              control={<Radio size="small" />}
              label={<Typography variant="caption">Fuzzy</Typography>}
            />
            <FormControlLabel
              value="regex"
              control={<Radio size="small" />}
              label={<Typography variant="caption">Regex</Typography>}
            />
          </RadioGroup>
        </Box>
      </Collapse>
    </Box>
  );
}
