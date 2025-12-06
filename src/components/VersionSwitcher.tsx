import { useState, useEffect } from "react";
import {
  IconButton,
  Menu,
  MenuItem,
  ListItemText,
  ListItemIcon,
  Typography,
  Divider,
  Chip,
  Tooltip,
  CircularProgress,
} from "@mui/material";
import HistoryIcon from "@mui/icons-material/History";
import CheckIcon from "@mui/icons-material/Check";
import OpenInNewIcon from "@mui/icons-material/OpenInNew";

interface VersionEntry {
  version: string;
  description: string;
  date: string;
}

interface VersionsManifest {
  latest: string;
  versions: VersionEntry[];
}

// Get the current version from meta tag (injected by shell) or URL
function getCurrentVersion(): string | null {
  // Check for meta tag first (set by the shell loader)
  const metaTag = document.querySelector('meta[name="app-version"]');
  if (metaTag) {
    return metaTag.getAttribute("content");
  }

  // Fallback: check URL query param
  const urlParams = new URLSearchParams(window.location.search);
  const vParam = urlParams.get("v");
  if (vParam) return vParam;

  // Fallback: check URL path pattern
  const match = window.location.pathname.match(/\/(v\d+(?:\.\d+)*)\/?/);
  return match ? match[1] : null;
}

// Check if we're running inside the shell's iframe
function isInShellIframe(): boolean {
  try {
    return window.parent !== window && window.parent.location.origin === window.location.origin;
  } catch {
    // Cross-origin - we're in an iframe but can't access parent
    return true;
  }
}

// Get the shell's base URL for fetching versions.json
function getShellBaseUrl(): string {
  // If in iframe, try to get parent's origin
  try {
    if (window.parent !== window) {
      return window.parent.location.origin + window.parent.location.pathname.replace(/\/[^/]*$/, "");
    }
  } catch {
    // Cross-origin, can't access parent
  }

  // Fallback: construct from current location
  // jsDelivr URL pattern: cdn.jsdelivr.net/gh/user/repo@version/dist/
  if (window.location.hostname === "cdn.jsdelivr.net") {
    // We're loaded directly from CDN, versions.json is on GitHub Pages
    // Extract user/repo from path
    const match = window.location.pathname.match(/\/gh\/([^@]+)@/);
    if (match) {
      const [user, repo] = match[1].split("/");
      return `https://${user}.github.io/${repo}`;
    }
  }

  // Default: assume versions.json is at the site root
  return window.location.origin;
}

export function VersionSwitcher() {
  const [anchorEl, setAnchorEl] = useState<null | HTMLElement>(null);
  const [versions, setVersions] = useState<VersionsManifest | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const currentVersion = getCurrentVersion();
  const open = Boolean(anchorEl);

  const handleClick = (event: React.MouseEvent<HTMLElement>) => {
    setAnchorEl(event.currentTarget);
    if (!versions && !loading) {
      loadVersions();
    }
  };

  const handleClose = () => {
    setAnchorEl(null);
  };

  const loadVersions = async () => {
    setLoading(true);
    setError(null);
    try {
      const baseUrl = getShellBaseUrl();
      const versionsUrl = `${baseUrl}/versions.json`;
      const response = await fetch(versionsUrl);
      if (!response.ok) {
        throw new Error("versions.json not found");
      }
      const data = await response.json();
      setVersions(data);
    } catch (err) {
      setError("Could not load version list");
      console.warn("Failed to load versions:", err);
    } finally {
      setLoading(false);
    }
  };

  const navigateToVersion = (version: VersionEntry) => {
    const baseUrl = getShellBaseUrl();

    if (isInShellIframe()) {
      // Navigate the parent window (the shell)
      try {
        window.parent.location.href = `${baseUrl}/?v=${version.version}`;
        return;
      } catch {
        // Cross-origin, fall through to regular navigation
      }
    }

    // Direct navigation (not in iframe or cross-origin)
    window.location.href = `${baseUrl}/?v=${version.version}`;
  };

  const openVersionsPage = () => {
    const baseUrl = getShellBaseUrl();

    if (isInShellIframe()) {
      try {
        window.parent.open(`${baseUrl}/versions.html`, "_blank");
        return;
      } catch {
        // Cross-origin, fall through
      }
    }

    window.open(`${baseUrl}/versions.html`, "_blank");
  };

  // Pre-load versions on mount for faster menu display
  useEffect(() => {
    loadVersions();
  }, []);

  // Don't render if we're not in a versioned deployment
  // (i.e., in development or non-versioned production)
  if (!currentVersion && !versions) {
    return null;
  }

  return (
    <>
      <Tooltip title="Switch version">
        <IconButton
          onClick={handleClick}
          color="inherit"
          size="small"
          sx={{ ml: 1 }}
        >
          <HistoryIcon />
          {currentVersion && (
            <Chip
              label={currentVersion}
              size="small"
              sx={{
                ml: 0.5,
                height: 20,
                fontSize: "0.75rem",
                backgroundColor: "rgba(144, 202, 249, 0.2)",
              }}
            />
          )}
        </IconButton>
      </Tooltip>
      <Menu
        anchorEl={anchorEl}
        open={open}
        onClose={handleClose}
        anchorOrigin={{
          vertical: "bottom",
          horizontal: "right",
        }}
        transformOrigin={{
          vertical: "top",
          horizontal: "right",
        }}
        PaperProps={{
          sx: { minWidth: 250, maxHeight: 400 },
        }}
      >
        <MenuItem disabled>
          <Typography variant="subtitle2" color="text.secondary">
            Available Versions
          </Typography>
        </MenuItem>
        <Divider />
        {loading && (
          <MenuItem disabled>
            <CircularProgress size={20} sx={{ mr: 2 }} />
            Loading versions...
          </MenuItem>
        )}
        {error && (
          <MenuItem disabled>
            <Typography color="error" variant="body2">
              {error}
            </Typography>
          </MenuItem>
        )}
        {versions?.versions.map((v) => {
          const isCurrent = v.version === currentVersion;
          const isLatest = v.version === versions.latest;
          return (
            <MenuItem
              key={v.version}
              onClick={() => navigateToVersion(v)}
              selected={isCurrent}
            >
              <ListItemIcon>
                {isCurrent ? <CheckIcon fontSize="small" /> : null}
              </ListItemIcon>
              <ListItemText
                primary={
                  <>
                    {v.version}
                    {isLatest && (
                      <Chip
                        label="Latest"
                        size="small"
                        color="success"
                        sx={{ ml: 1, height: 18, fontSize: "0.7rem" }}
                      />
                    )}
                  </>
                }
                secondary={
                  <>
                    {v.description}
                    <br />
                    <Typography variant="caption" color="text.disabled">
                      {v.date}
                    </Typography>
                  </>
                }
              />
            </MenuItem>
          );
        })}
        {versions && (
          <>
            <Divider />
            <MenuItem onClick={openVersionsPage}>
              <ListItemIcon>
                <OpenInNewIcon fontSize="small" />
              </ListItemIcon>
              <ListItemText primary="View all versions" />
            </MenuItem>
          </>
        )}
      </Menu>
    </>
  );
}
