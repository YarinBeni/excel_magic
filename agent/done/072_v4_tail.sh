# TIMEOUT=120
f=$(ls -t sbatch/logs/V4_rel_capability_51838*.out | head -1); echo "== $f"; tail -c 3000 "$f"
e=$(ls -t sbatch/logs/V4_rel_capability_51838*.err 2>/dev/null | head -1); [ -n "$e" ] && { echo "== $e"; tail -c 2500 "$e"; }
srun --jobid 51838 --overlap nvidia-smi --query-gpu=utilization.gpu,memory.used --format=csv 2>&1 | tail -3
srun --jobid 51838 --overlap ps -u "$USER" -o pid,etime,pcpu,args 2>&1 | grep -E "python" | grep -v grep | cut -c1-160
d=$(ls -td artifacts/runs/*V6_insightbench_pi_all | head -1); python3 - "$d" <<'P'
import json,sys
for c in "DEF":
    try: rows=[json.loads(l) for l in open(f"{sys.argv[1]}/insights_pi{c}.jsonl")]
    except FileNotFoundError: continue
    err=sum(1 for r in rows if r["error"]); zero=sum(1 for r in rows if not r["insights"])
    print(c, len(rows), "err", err, "zero", zero, "ins", round(sum(len(r["insights"]) for r in rows)/len(rows),2), "rej", round(sum(r["n_rejected"] for r in rows)/len(rows),2))
    print("  ", [r["error"][:120] for r in rows if r["error"]][:2])
P
