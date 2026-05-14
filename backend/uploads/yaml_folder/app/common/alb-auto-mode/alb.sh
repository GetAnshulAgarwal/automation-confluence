#!/bin/bash -x

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

kubectl config set-context $CONTEXT
kubectl apply -f "$SCRIPT_DIR/ingressClassParams.yaml"
kubectl apply -f "$SCRIPT_DIR/ingressClass.yaml"

# Test
kubectl get ingressclassparams alb-internal-app -o yaml
kubectl get ingressclass alb-internal-app -o yaml