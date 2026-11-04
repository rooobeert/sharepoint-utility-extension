#!/bin/sh
# Builds the store package: dist/sharepoint-utility-<version>.zip with
# manifest.json at its root. Runs the tests and checks the manifest first.
# Usage: npm run package
set -eu
cd "$(dirname "$0")/.."

node --test tests/*.test.js >/dev/null || { echo "Tests failed; run npm test for details." >&2; exit 1; }

node --input-type=module -e '
import { readFileSync, existsSync } from "node:fs";
const m = JSON.parse(readFileSync("extension/manifest.json", "utf8"));
const errors = [];
if (m.manifest_version !== 3) errors.push("manifest_version must be 3");
if (!/^\d+(\.\d+){0,3}$/.test(m.version)) errors.push(`version "${m.version}" is not 1-4 dot-separated numbers`);
if (m.description.length > 132) errors.push(`description is ${m.description.length} characters; the stores allow 132`);
const icons = { ...(m.icons || {}), ...((m.action && m.action.default_icon) || {}) };
if (!m.icons || !m.icons["128"]) errors.push("icons.128 is missing; both stores require a 128x128 icon");
for (const p of new Set(Object.values(icons))) if (!existsSync(`extension/${p}`)) errors.push(`icon file extension/${p} does not exist`);
if (errors.length) { console.error("Manifest problems:\n- " + errors.join("\n- ")); process.exit(1); }
'

VERSION=$(node -p 'JSON.parse(require("fs").readFileSync("extension/manifest.json","utf8")).version')
ZIP="dist/sharepoint-utility-$VERSION.zip"
mkdir -p dist
rm -f "$ZIP"
(cd extension && zip -qr -X "../$ZIP" . -x '.*' -x '*/.*')
echo "Built $ZIP ($(du -h "$ZIP" | cut -f1 | tr -d ' '), $(unzip -Z1 "$ZIP" | wc -l | tr -d ' ') files)"
