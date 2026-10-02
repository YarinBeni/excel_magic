#!/usr/bin/env bash
# Downloads TabPFN v2 weights from the public GCS mirror (no Hugging Face access needed).
set -euo pipefail
cd "$(dirname "$0")/.."
mkdir -p weights/tabpfn
BASE="https://storage.googleapis.com/tabpfn-v2-model-files/05152025"
for f in tabpfn-v2-classifier.ckpt tabpfn-v2-regressor.ckpt; do
  if [ ! -s "weights/tabpfn/$f" ]; then
    echo "downloading $f"; curl -sS -L --fail -o "weights/tabpfn/$f" "$BASE/$f"
  fi
done
ls -la weights/tabpfn
# TabICL v2 / OpenRFM are optional; see docs/BLOCKERS.md for where their weights live.
