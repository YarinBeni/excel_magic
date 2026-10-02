# _vllm.sh — start a vLLM server for $LLM on $PORT with tool calling, trying parsers in order; sets VLLM_PID.
#   LLM=Qwen/Qwen3-Coder-30B-A3B-Instruct PORT=8000 start_vllm "qwen3_coder hermes"
# vLLM lives in its OWN conda env (python 3.12): a venv on top of the thesis python mixed conda's libicu with the
# system libstdc++ (CXXABI_1.3.15 not found when vllm imported sqlite3).
VLLM_VENV="$HOME/miniconda3/envs/vllm"
ensure_vllm() {
    # one installer at a time (several GPU jobs start together): flock + an .ok marker written only after `vllm --version`
    ( flock -w 3600 9 || { echo "[vllm] could not get install lock"; exit 1; }
      if [ -f "$VLLM_VENV/.ok" ] && [ -x "$VLLM_VENV/bin/vllm" ]; then exit 0; fi
      echo "[vllm] installing into $VLLM_VENV (one-time, under lock)"
      source ~/miniconda3/etc/profile.d/conda.sh
      conda env remove -y -n vllm >/dev/null 2>&1 || true
      conda create -y -q -n vllm python=3.12 pip || exit 1
      "$VLLM_VENV/bin/pip" install -q --upgrade pip || exit 1
      "$VLLM_VENV/bin/pip" install -q --no-cache-dir vllm || exit 1
      "$VLLM_VENV/bin/vllm" --version || exit 1
      touch "$VLLM_VENV/.ok"
    ) 9>"$HOME/.vllm_install.lock" || { echo "FAILED vllm install"; return 1; }
}
start_vllm() {
    local parsers="$1"; local extra="${2:-}"; local p
    ensure_vllm || return 1
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
        echo "[vllm] failed with parser $p:"; tail -15 "artifacts/vllm/${SLURM_JOB_NAME:-job}_${SLURM_JOB_ID:-0}_$p.log"
    done
    return 1
}
