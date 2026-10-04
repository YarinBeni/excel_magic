# TIMEOUT=60
R=$(ls -d artifacts/runs/20261004T104448Z_V7_sql_harness2)
RUN="$R" MODEL=Qwen/Qwen3-Coder-30B-A3B-Instruct TAG=qwen3coder30b CONFIGS=A,SC8N,GN,GN8 SPLIT=heldout sbatch sbatch/V7_sql_harness2.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %.12l %R"
