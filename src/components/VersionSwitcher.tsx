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
  path: string;
}

interface VersionsManifest {
  latest: string;
  versions: VersionEntry[];
}

// Get the current version from the URL path
function getCurrentVersion(): string | null {
  const match = window.location.pathname.match(/\/(v\d+(?:\.\d+)*)\//);
  return match ? match[1] : null;
}

// Get the base path for versions (everything before the version)
function getBasePath(): string {
  const path = window.location.pathname;
  const match = path.match(/(.*)\/v\d+(?:\.\d+)*\//);
  if (match) {
    return match[1];
  }
  // If no version in path, return the current path without trailing slash
  return path.replace(/\/$/, "");
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
      // Try to load versions.json from the root of the GitHub Pages site
      const basePath = getBasePath();
      const versionsUrl = `${basePath}/versions.json`;
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
    const basePath = getBasePath();
    window.location.href = `${basePath}${version.path}`;
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
            <MenuItem
              onClick={() => {
                const basePath = getBasePath();
                window.open(`${basePath}/versions.html`, "_blank");
              }}
            >
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
