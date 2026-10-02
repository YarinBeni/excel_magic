# _vllm.sh — start a vLLM server for $LLM on $PORT with tool calling, trying parsers in order; sets VLLM_PID.
#   LLM=Qwen/Qwen3-Coder-30B-A3B-Instruct PORT=8000 start_vllm "qwen3_coder hermes"
# vLLM lives in its OWN conda env (python 3.12): a venv on top of the thesis python mixed conda's libicu with the
# system libstdc++ (CXXABI_1.3.15 not found when vllm imported sqlite3).
VLLM_VENV="$HOME/miniconda3/envs/vllm"
# conda's libicu needs conda's (newer) libstdc++; without this the system libstdc++ gets loaded first and
# `import sqlite3` dies with CXXABI_1.3.15 not found.
vllm_env() { export LD_LIBRARY_PATH="$VLLM_VENV/lib:${LD_LIBRARY_PATH:-}"; }
ensure_vllm() {
    # one installer at a time (several GPU jobs start together): flock + an .ok marker written only after `vllm --version`
    ( flock -w 3600 9 || { echo "[vllm] could not get install lock"; exit 1; }
      if [ -f "$VLLM_VENV/.ok" ] && [ -x "$VLLM_VENV/bin/vllm" ]; then exit 0; fi
      echo "[vllm] installing into $VLLM_VENV (one-time, under lock)"
      source ~/miniconda3/etc/profile.d/conda.sh
      conda env remove -y -n vllm >/dev/null 2>&1 || true
      conda create -y -q -n vllm python=3.12 pip || exit 1
      "$VLLM_VENV/bin/pip" install -q --upgrade pip || exit 1
      vllm_env
      # The cluster driver is CUDA 12.8 (torch reports "driver too old" for cu130 wheels, which vllm >= 0.18 pulls via torch 2.10+).
      # Pin a vLLM whose torch is a CUDA 12 build (0.16.0 -> torch 2.9.1+cu128); fall back to 0.11.0 (torch 2.8.0+cu128).
      # The check runs on the CPU install node, so it asserts the CUDA build, not a device: not `vllm --version` either (that
      # builds the CLI parser, which infers a device).
      ok=""
      for pin in "${VLLM_PIN:-0.16.0}" 0.11.0; do
        echo "[vllm] pip install vllm==$pin"
        "$VLLM_VENV/bin/pip" install -q --no-cache-dir "vllm==$pin" || continue
        "$VLLM_VENV/bin/python" -c "import torch, vllm; c=torch.version.cuda; assert c and c.startswith('12.'), f'torch CUDA {c} is not a 12.x build'; print('vllm', vllm.__version__, 'torch', torch.__version__, 'cuda', c)" && { ok=1; break; }
      done
      [ -n "$ok" ] || exit 1
      touch "$VLLM_VENV/.ok"
    ) 9>"$HOME/.vllm_install.lock" || { echo "FAILED vllm install"; return 1; }
}
start_vllm() {
    local parsers="$1"; local extra="${2:-}"; local p
    ensure_vllm || return 1
    vllm_env
    mkdir -p artifacts/vllm
    for p in $parsers; do
        echo "[vllm] serving $LLM on :$PORT with --tool-call-parser $p"
        "$VLLM_VENV/bin/vllm" serve "$LLM" --port "$PORT" --enable-auto-tool-choice --tool-call-parser "$p" $extra \
            --max-model-len 32768 --gpu-memory-utilization "${VLLM_GPU_FRAC:-0.70}" > "artifacts/vllm/${SLURM_JOB_NAME:-job}_${SLURM_JOB_ID:-0}_$p.log" 2>&1 &
        VLLM_PID=$!
        local i
        for i in $(seq 1 160); do
            curl -sf "http://localhost:$PORT/v1/models" >/dev/null 2>&1 && { echo "[vllm] up (parser $p)"; return 0; }
            kill -0 "$VLLM_PID" 2>/dev/null || break
            sleep 15
        done
        kill "$VLLM_PID" 2>/dev/null || true
        local lg="artifacts/vllm/${SLURM_JOB_NAME:-job}_${SLURM_JOB_ID:-0}_$p.log"
        echo "[vllm] failed with parser $p; root-cause lines from $lg:"
        # the APIServer traceback only says "see root cause above": show the engine-side errors instead
        grep -aiE "error|exception|out of memory|memory|not supported|unrecognized|no such|not found|assert" "$lg" \
            | grep -av "APIServer pid" | grep -av "^\s*File " | tail -14
        tail -3 "$lg"
    done
    return 1
}
