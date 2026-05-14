#!/bin/bash
set -o pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [[ -z "${CONTEXT:-}" ]]; then
  echo "Error: CONTEXT is not set. Export CONTEXT before running this script." >&2
  exit 1
fi

kubectl --context "$CONTEXT" create configmap keycloak-realm \
  --from-file=realm.json="$SCRIPT_DIR/lio-realm-v1.json" \
  -o yaml --dry-run=client \
  | kubectl --context "$CONTEXT" -n keycloak apply -f -