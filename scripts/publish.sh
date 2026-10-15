#!/usr/bin/env bash
#
# Publish the latest product data to the site.
#
# Partners collaborate on the master spreadsheet in OneDrive:
#   ~/OneDrive/Stride-Website/product-data.xlsx
# This script copies that master into the repo, regenerates every page from
# it, and shows what changed. Review, then commit & push to go live.
#
# Override the master location if needed:  STRIDE_DATA=/path/to.xlsx scripts/publish.sh
#
set -euo pipefail
cd "$(dirname "$0")/.."

MASTER="${STRIDE_DATA:-$HOME/Library/CloudStorage/OneDrive-Personal/Stride-Website/product-data.xlsx}"

if [ -f "$MASTER" ]; then
  echo "→ Syncing product data from OneDrive:"
  echo "    $MASTER"
  cp "$MASTER" product-data.xlsx
else
  echo "⚠  OneDrive master not found at:"
  echo "    $MASTER"
  echo "   Falling back to the repo's committed product-data.xlsx."
fi

echo "→ Regenerating site..."
python3 scripts/generate_products.py

echo
echo "→ Changes:"
git status --short || true
echo
echo "Review the changes above, then publish:"
echo "    git add -A && git commit -m \"Update catalog\" && git push        # to staging"
echo "Ask Claude (or merge staging → main) when you're ready to go live."
