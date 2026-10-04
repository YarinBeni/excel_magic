# TIMEOUT=60
MODEL=Qwen/Qwen3-Coder-30B-A3B-Instruct TAG=qwen3coder30b CONFIGS=A,GN,GNL,GNLz SPLIT=heldout sbatch sbatch/V7_sql_harness2.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %.12l %R"
