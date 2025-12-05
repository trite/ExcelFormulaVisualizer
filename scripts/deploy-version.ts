#!/usr/bin/env npx tsx
/**
 * Deployment script for versioned GitHub Pages releases
 *
 * Usage:
 *   npx tsx scripts/deploy-version.ts <version> [description]
 *
 * Examples:
 *   npx tsx scripts/deploy-version.ts v1 "Initial release"
 *   npx tsx scripts/deploy-version.ts v2 "Added node filtering"
 *
 * This script:
 * 1. Builds the app with the correct base path for the version
 * 2. Updates versions.json manifest
 * 3. Prepares the dist folder for deployment
 */

import { execSync } from "child_process";
import * as fs from "fs";
import * as path from "path";

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

const DIST_DIR = "dist";
const VERSIONS_FILE = "versions.json";

function loadVersionsManifest(): VersionsManifest {
  const manifestPath = path.join(DIST_DIR, VERSIONS_FILE);
  if (fs.existsSync(manifestPath)) {
    return JSON.parse(fs.readFileSync(manifestPath, "utf-8"));
  }
  return { latest: "", versions: [] };
}

function saveVersionsManifest(manifest: VersionsManifest): void {
  const manifestPath = path.join(DIST_DIR, VERSIONS_FILE);
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));
}

function main() {
  const args = process.argv.slice(2);

  if (args.length < 1) {
    console.error("Usage: npx tsx scripts/deploy-version.ts <version> [description]");
    console.error("Example: npx tsx scripts/deploy-version.ts v2 'Added node filtering'");
    process.exit(1);
  }

  const version = args[0];
  const description = args[1] || `Version ${version}`;

  // Validate version format (supports v0.1, v1, v1.2.3, etc.)
  if (!/^v\d+(\.\d+)*$/.test(version)) {
    console.error(`Invalid version format: ${version}`);
    console.error("Version should be like: v0.1, v1, v1.2, v1.2.3, etc.");
    process.exit(1);
  }

  console.log(`\n📦 Deploying version: ${version}`);
  console.log(`📝 Description: ${description}\n`);

  // Step 1: Build with versioned base path
  const basePath = `/ExcelFormulaVisualizer/${version}/`;
  console.log(`🔨 Building with base path: ${basePath}`);

  // Set environment variable for vite config to pick up
  process.env.VITE_VERSION = version;
  process.env.VITE_BASE_PATH = basePath;

  execSync(`npx vite build --base=${basePath}`, {
    stdio: "inherit",
    env: { ...process.env }
  });

  // Step 2: Create versioned directory structure
  const versionDir = path.join(DIST_DIR, version);
  const tempDir = path.join(DIST_DIR, "_temp_build");

  // Move current build to temp
  if (fs.existsSync(tempDir)) {
    fs.rmSync(tempDir, { recursive: true });
  }

  // Move all built files except versions.json to temp
  const builtFiles = fs.readdirSync(DIST_DIR).filter(f => f !== VERSIONS_FILE && f !== version);
  fs.mkdirSync(tempDir, { recursive: true });

  for (const file of builtFiles) {
    const src = path.join(DIST_DIR, file);
    const dest = path.join(tempDir, file);
    fs.renameSync(src, dest);
  }

  // Create version directory and move files there
  if (fs.existsSync(versionDir)) {
    fs.rmSync(versionDir, { recursive: true });
  }
  fs.renameSync(tempDir, versionDir);

  console.log(`✅ Built files moved to ${versionDir}`);

  // Step 3: Update versions manifest
  const manifest = loadVersionsManifest();

  // Remove existing entry for this version if it exists
  manifest.versions = manifest.versions.filter(v => v.version !== version);

  // Add new version entry
  manifest.versions.unshift({
    version,
    description,
    date: new Date().toISOString().split("T")[0],
    path: `/${version}/`,
  });

  // Sort versions (newest first) using semantic versioning comparison
  manifest.versions.sort((a, b) => {
    const aParts = a.version.replace(/^v/, "").split(".").map(Number);
    const bParts = b.version.replace(/^v/, "").split(".").map(Number);
    for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
      const aVal = aParts[i] || 0;
      const bVal = bParts[i] || 0;
      if (bVal !== aVal) return bVal - aVal;
    }
    return 0;
  });

  manifest.latest = manifest.versions[0].version;
  saveVersionsManifest(manifest);

  console.log(`✅ Updated ${VERSIONS_FILE}`);

  // Step 4: Create root index.html that redirects to latest
  const rootIndexHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="refresh" content="0;url=./${manifest.latest}/">
  <title>Excel Formula Visualizer</title>
  <style>
    body {
      font-family: system-ui, sans-serif;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      margin: 0;
      background: #121212;
      color: #fff;
    }
  </style>
</head>
<body>
  <p>Redirecting to latest version (${manifest.latest})...</p>
  <script>window.location.href = "./${manifest.latest}/";</script>
</body>
</html>`;

  fs.writeFileSync(path.join(DIST_DIR, "index.html"), rootIndexHtml);
  console.log(`✅ Created root index.html with redirect to ${manifest.latest}`);

  // Step 5: Create a simple version listing page
  const versionsHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Excel Formula Visualizer - All Versions</title>
  <style>
    body {
      font-family: system-ui, sans-serif;
      max-width: 600px;
      margin: 40px auto;
      padding: 20px;
      background: #121212;
      color: #fff;
    }
    h1 { color: #90caf9; }
    .version {
      padding: 15px;
      margin: 10px 0;
      background: #1e1e1e;
      border-radius: 8px;
      border-left: 4px solid #90caf9;
    }
    .version.latest { border-left-color: #4caf50; }
    .version a {
      color: #90caf9;
      text-decoration: none;
      font-size: 1.2em;
      font-weight: bold;
    }
    .version a:hover { text-decoration: underline; }
    .version .meta {
      color: #888;
      font-size: 0.9em;
      margin-top: 5px;
    }
    .version .desc { margin-top: 8px; }
    .badge {
      display: inline-block;
      background: #4caf50;
      color: #fff;
      padding: 2px 8px;
      border-radius: 4px;
      font-size: 0.8em;
      margin-left: 10px;
    }
  </style>
</head>
<body>
  <h1>Excel Formula Visualizer</h1>
  <p>Select a version to use:</p>
  <div id="versions"></div>
  <script>
    fetch('./versions.json')
      .then(r => r.json())
      .then(data => {
        const container = document.getElementById('versions');
        data.versions.forEach(v => {
          const isLatest = v.version === data.latest;
          container.innerHTML += \`
            <div class="version \${isLatest ? 'latest' : ''}">
              <a href=".\${v.path}">\${v.version}</a>
              \${isLatest ? '<span class="badge">Latest</span>' : ''}
              <div class="meta">\${v.date}</div>
              <div class="desc">\${v.description}</div>
            </div>
          \`;
        });
      });
  </script>
</body>
</html>`;

  fs.writeFileSync(path.join(DIST_DIR, "versions.html"), versionsHtml);
  console.log(`✅ Created versions.html listing page`);

  console.log(`\n🎉 Version ${version} is ready for deployment!`);
  console.log(`\nTo deploy, run:`);
  console.log(`  gh-pages -d dist --add\n`);
  console.log(`The --add flag preserves existing versions on GitHub Pages.\n`);
}

main();
