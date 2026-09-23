#!/bin/bash
# SubCourt — clasp deploy script
# Usage:
#   ./deploy.sh prod    — push to PROD Apps Script

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

PROD_DEPLOYMENT_ID="AKfycbzb3EnQsxBt5dLTaQpg7VJjtoBtHTyGpB2VgpfJ9TDuvezk0ihjhn5oW48a9oKiIAyYMg"
PROD_SHEET_ID="1hA-ZPhV62pp376qtWRDfQQkFv6y9U5Wkm0nUyKCHC6o"

TARGET="${1:-}"
DESCRIPTION="${2:-}"

if [ "$TARGET" != "prod" ]; then
  echo "Usage: ./deploy.sh prod [description]"
  exit 1
fi

echo "→ Deploying to PROD..."
cp "$SCRIPT_DIR/SubCourt-AppScript-PROD.js" "$SCRIPT_DIR/clasp/prod/Code.js"
cp "$SCRIPT_DIR/appsscript.json" "$SCRIPT_DIR/clasp/prod/appsscript.json"
cd "$SCRIPT_DIR/clasp/prod"
clasp push --force
if [ -n "$DESCRIPTION" ]; then
  clasp deploy --deploymentId "$PROD_DEPLOYMENT_ID" --description "$DESCRIPTION"
else
  clasp deploy --deploymentId "$PROD_DEPLOYMENT_ID"
fi
echo "✓ PROD deploy complete and live."
