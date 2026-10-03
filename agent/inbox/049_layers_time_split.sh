# TIMEOUT=300
# Protocol fix. With a random split of context vs probe-train rows, probe-train rows share periods and entities with
# the context; under a real-label context the late layers carry a copy of a neighbour's label that val/test rows never
# get, so the linear probe fails on val/test (J21 rel-f1 late layers: linear val 0.44-0.55, kNN ~0.70). All three
# studies now default to a time split (probe-train rows strictly later than the context, like val/test). Rerun all
# tasks, rel-f1 included (the first-pass rel-f1 runs stay as the random-split comparison).
for J in J18 J19 J21; do id=$(cat agent/state/$J.id 2>/dev/null || true); [ -n "$id" ] && scancel "$id" && echo "cancelled $J $id"; done
sleep 5
J18=$(sbatch --parsable --time=12:00:00 sbatch/J18_relbench_layers.sbatch); echo "$J18" > agent/state/J18.id; echo "J18: $J18"
J19=$(sbatch --parsable --time=12:00:00 sbatch/J19_relbench_tab_layers.sbatch); echo "$J19" > agent/state/J19.id; echo "J19: $J19"
J21=$(sbatch --parsable --time=14:00:00 sbatch/J21_model_layers.sbatch); echo "$J21" > agent/state/J21.id; echo "J21: $J21"
grep -n "ctx_split\|def time_split" frozen-embeddings-retrieval/fer/relbench_layers.py | head -3
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %.10l %R" | sort -u
