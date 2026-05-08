#!/bin/bash
# SubCourt — clasp deploy script
# Usage:
#   ./deploy.sh dev     — push to DEV Apps Script (rally-tennis-dev.html)
#   ./deploy.sh prod    — push to PROD Apps Script

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SOURCE="$SCRIPT_DIR/SubCourt-AppScript.js"

DEV_DEPLOYMENT_ID="AKfycbxS8vYTuuuxsjbVoLS0Mup8VYiCj0t95N6dq7cCKIimnwfLW4or5qBoGFHGbVZIT597Ug"
PROD_DEPLOYMENT_ID="AKfycbymmYjtMAEnNO0D_1yuYZJQmOisxtSHBcaOCzTm74-iBmSnQgQgFTMB4IvJqeK6SEFhMg"
DEV_SHEET_ID="1VjFuq63KLEgZpYvCVi2bJrWEgMxDP6hXygYwjDpUmRE"
PROD_SHEET_ID="1hA-ZPhV62pp376qtWRDfQQkFv6y9U5Wkm0nUyKCHC6o"

TARGET="${1:-}"
DESCRIPTION="${2:-}"

if [ -z "$TARGET" ]; then
  echo "Usage: ./deploy.sh dev | prod [description]"
  exit 1
fi

if [ "$TARGET" = "dev" ]; then
  echo "→ Deploying to DEV..."
  cp "$SOURCE" "$HOME/subcourt-dev-script/SubCourt-AppScript.js"
  cd "$HOME/subcourt-dev-script"
  clasp push --force
  if [ -n "$DESCRIPTION" ]; then
    clasp deploy --deploymentId "$DEV_DEPLOYMENT_ID" --description "$DESCRIPTION"
  else
    clasp deploy --deploymentId "$DEV_DEPLOYMENT_ID"
  fi
  echo "✓ DEV deploy complete and live."

elif [ "$TARGET" = "prod" ]; then
  if [ -z "$PROD_SHEET_ID" ]; then
    echo "Error: PROD_SHEET_ID is not set in deploy.sh. Add it before deploying to prod."
    exit 1
  fi
  echo "→ Deploying to PROD..."
  sed "s/$DEV_SHEET_ID/$PROD_SHEET_ID/;s/rally-tennis-dev\.html/rally-tennis-prod.html/g" "$SOURCE" > "$SCRIPT_DIR/clasp/prod/Code.js"
  cd "$SCRIPT_DIR/clasp/prod"
  clasp push --force
  if [ -n "$DESCRIPTION" ]; then
    clasp deploy --deploymentId "$PROD_DEPLOYMENT_ID" --description "$DESCRIPTION"
  else
    clasp deploy --deploymentId "$PROD_DEPLOYMENT_ID"
  fi
  echo "✓ PROD deploy complete and live."

else
  echo "Unknown target: $TARGET. Use 'dev' or 'prod'."
  exit 1
fi
