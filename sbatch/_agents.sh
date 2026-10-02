# _agents.sh — helpers for the free CLI coding agents. setup_pi BASE LLM: node22 env on PATH, pi installed, provider config.
setup_pi() {
    local base="$1" llm="$2"
    ( flock -w 1800 8 || exit 1
      [ -x "$HOME/miniconda3/envs/node22/bin/node" ] || conda create -y -q -n node22 -c conda-forge "nodejs>=22" || exit 1
    ) 8>"$HOME/.node22_install.lock" || { echo "FAILED conda node22"; return 1; }
    export PATH="$HOME/miniconda3/envs/node22/bin:$PATH"
    command -v pi >/dev/null || npm i -g --ignore-scripts @earendil-works/pi-coding-agent || { echo "FAILED npm pi"; return 1; }
    mkdir -p ~/.pi/agent
    printf '{"providers":{"vllm":{"baseUrl":"%s","api":"openai-completions","apiKey":"dummy","models":[{"id":"%s"}]}}}\n' "$base" "$llm" > ~/.pi/agent/models.json
}
