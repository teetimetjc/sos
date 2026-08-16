#!/usr/bin/env bash
# Push latest Code.gs + App.html to Apps Script and create a new deployment version.
# Usage: ./deploy.sh
# Requires: clasp installed (npm install -g @google/clasp) and logged in (clasp login)

set -euo pipefail

echo "→ Pushing files to Apps Script…"
clasp push --force

echo "→ Creating new deployment version…"
clasp deploy --description "$(date '+%Y-%m-%d %H:%M') deploy"

echo "✓ Done. Open the Apps Script project to grab the new /exec URL if needed:"
clasp open
