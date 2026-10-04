# TIMEOUT=60
MODEL=Qwen/Qwen3-Coder-30B-A3B-Instruct TAG=qwen3coder30b CONFIGS=A,K1,K2,K3,G,GD SPLIT=dev sbatch sbatch/V7_sql_harness2.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R"
