# TIMEOUT=120
RUN=$(ls -td artifacts/runs/*V4_relbench_capability | head -1); f="$RUN/capability_rows.jsonl"; echo "$RUN"
python3 - "$f" <<'P'
import json,sys
rows=[json.loads(l) for l in open(sys.argv[1])]
keep=[r for r in rows if not (r["task"]=="rel-hm/user-churn" and r["config"] in ("E","ER"))]
open(sys.argv[1],"w").write("".join(json.dumps(r)+"\n" for r in keep)); print(len(rows),"->",len(keep))
P
RUN="$RUN" TASKS=rel-hm/user-churn CONFIGS=E,ER sbatch sbatch/V4_relbench_capability.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R"
