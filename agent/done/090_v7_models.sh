# TIMEOUT=60
MODEL=openai/gpt-oss-120b TAG=gptoss120b PARSERS=openai VLLM_GPU_FRAC=0.90 CONFIGS=A,K3,G SPLIT=dev sbatch sbatch/V7_sql_harness2.sbatch
MODEL=zai-org/GLM-4.5-Air-FP8 TAG=glm45air PARSERS="glm45 hermes" VLLM_GPU_FRAC=0.90 CONFIGS=A,K3,G SPLIT=dev sbatch sbatch/V7_sql_harness2.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R"
