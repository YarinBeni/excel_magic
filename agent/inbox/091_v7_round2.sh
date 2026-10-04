# TIMEOUT=60
scancel 52319 52320 2>/dev/null; sleep 2
R=$(ls -td artifacts/runs/*V7_sql_harness2 | head -1); echo "qwen run: $R"
RUN="$R" MODEL=Qwen/Qwen3-Coder-30B-A3B-Instruct TAG=qwen3coder30b CONFIGS=GN,GN8,SC8N SPLIT=dev sbatch sbatch/V7_sql_harness2.sbatch
MODEL=openai/gpt-oss-120b TAG=gptoss120b PARSERS=openai VLLM_GPU_FRAC=0.90 CONFIGS=A,K2,GN SPLIT=dev sbatch sbatch/V7_sql_harness2.sbatch
MODEL=zai-org/GLM-4.5-Air-FP8 TAG=glm45air PARSERS="glm45 hermes" VLLM_GPU_FRAC=0.90 CONFIGS=A,K2,GN SPLIT=dev sbatch sbatch/V7_sql_harness2.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %.12l %R"
