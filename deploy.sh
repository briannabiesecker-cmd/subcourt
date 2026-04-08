#!/bin/bash
# SubCourt — clasp deploy script
# Usage:
#   ./deploy.sh test    — push to TEST Apps Script
#   ./deploy.sh prod    — push to PROD Apps Script

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SOURCE="$SCRIPT_DIR/SubCourt-AppScript.js"

TEST_SHEET_ID="1GLWl0a6lRgHsrpG5sZ3S8LtY7HJUGJplNCiPUHIuyIw"
PROD_SHEET_ID="1hA-ZPhV62pp376qtWRDfQQkFv6y9U5Wkm0nUyKCHC6o"

TARGET="${1:-}"

if [ -z "$TARGET" ]; then
  echo "Usage: ./deploy.sh test | prod"
  exit 1
fi

if [ "$TARGET" = "test" ]; then
  echo "→ Deploying to TEST..."
  cp "$SOURCE" "$SCRIPT_DIR/clasp/test/Code.js"
  cd "$SCRIPT_DIR/clasp/test"
  clasp push --force
  echo "✓ TEST deploy complete. Create a new version in the Apps Script editor (Deploy → Manage deployments) to make it live."

elif [ "$TARGET" = "prod" ]; then
  if [ -z "$PROD_SHEET_ID" ]; then
    echo "Error: PROD_SHEET_ID is not set in deploy.sh. Add it before deploying to prod."
    exit 1
  fi
  echo "→ Deploying to PROD..."
  sed "s/$TEST_SHEET_ID/$PROD_SHEET_ID/" "$SOURCE" > "$SCRIPT_DIR/clasp/prod/Code.js"
  cd "$SCRIPT_DIR/clasp/prod"
  clasp push --force
  echo "✓ PROD deploy complete. Create a new version in the Apps Script editor (Deploy → Manage deployments) to make it live."

else
  echo "Unknown target: $TARGET. Use 'test' or 'prod'."
  exit 1
fi
