# TIMEOUT=600
# J18/J19 sat ~2 h on rel-trial/study-outcome with no event: Subgrapher re-sorted each child table once per entity
# (quadratic). Fixed (sort once). Dump stacks to confirm, then restart J18/J19 on the 6 tasks left. J21's current
# task (rel-f1, tiny tables) is not affected and its next task picks up the fix; restart it only if it has no event
# for 40 min.
REST="rel-trial/study-outcome,rel-event/user-repeat,rel-event/user-ignore,rel-avito/user-visits,rel-avito/user-clicks,rel-hm/user-churn"
command -v py-spy >/dev/null || python -m pip install -q py-spy 2>&1 | tail -1
stack() {  # $1 = job id: show the python stack inside the job (needs py-spy and ptrace rights; best effort)
  timeout 90 srun --jobid="$1" --overlap -N1 -n1 bash -c 'for p in $(pgrep -u "$USER" -f "experiments/exp0"); do echo "pid $p $(ps -o etime=,pcpu=,rss= -p $p)"; command -v py-spy >/dev/null && py-spy dump --pid $p 2>&1 | grep -E "^Thread|Error|\(fer/|\(experiments/|relbench_|model_layers" | head -14; done' 2>&1 | head -40
}
for J in J18 J19 J21; do echo "== $J $(cat agent/state/$J.id)"; stack "$(cat agent/state/$J.id)" || true; done
d=$(ls -td artifacts/runs/*J21_model_layers_* 2>/dev/null | head -1); echo "== J21 run $d"
wc -l < "$d/events.jsonl"; tail -3 "$d/events.jsonl" | cut -c1-240
age=$(( $(date +%s) - $(stat -c %Y "$d/events.jsonl") )); echo "J21 last event ${age}s ago"
for J in J18 J19; do id=$(cat agent/state/$J.id); scancel "$id" && echo "cancelled $J $id"; done
if [ "$age" -gt 2400 ]; then id=$(cat agent/state/J21.id); scancel "$id" && echo "cancelled J21 $id (stalled)"; RESTART21=1; fi
sleep 5
J18=$(sbatch --parsable --time=10:00:00 --export=ALL,TASKS=$REST sbatch/J18_relbench_layers.sbatch); echo "$J18" > agent/state/J18.id; echo "J18: $J18"
J19=$(sbatch --parsable --time=10:00:00 --export=ALL,TASKS=$REST sbatch/J19_relbench_tab_layers.sbatch); echo "$J19" > agent/state/J19.id; echo "J19: $J19"
if [ -n "${RESTART21:-}" ]; then J21=$(sbatch --parsable sbatch/J21_model_layers.sbatch); echo "$J21" > agent/state/J21.id; echo "J21: $J21"; fi
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
