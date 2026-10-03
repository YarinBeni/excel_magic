# TIMEOUT=120
# No pushes from J18/J19/J21 for ~1.5 h: show where each job is.
for j in J18_relbench_layers J19_relbench_tab_layers J21_model_layers; do
  f=$(ls -t sbatch/logs/${j}_*.out 2>/dev/null | head -1); echo "== $f"
  grep -a -E "^--- |FAILED|Error|Traceback|embed_progress|layer|seconds" "$f" | tail -6 | cut -c1-220
  tail -2 "$f" | cut -c1-220
done
for d in $(ls -td artifacts/runs/*J18_layers* artifacts/runs/*J19_tab* artifacts/runs/*J21_model* 2>/dev/null | head -3); do echo "== $d"; tail -3 $d/events.jsonl 2>/dev/null | cut -c1-260; done
free -g | head -2; nvidia-smi --query-gpu=index,utilization.gpu,memory.used --format=csv 2>/dev/null | head -4
