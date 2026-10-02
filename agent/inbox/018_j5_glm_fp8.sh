# TIMEOUT=300
# J5 task 2 (GLM-4.5-Air) died at vLLM engine start: 106B-A12B in bf16 does not fit one H200. Now GLM-4.5-Air-FP8 at
# 0.90 GPU memory. Resubmit only array task 2; tasks 0/1/3 are untouched.
J5b=$(sbatch --parsable --array=2 sbatch/J5_llm_sweep.sbatch); echo "$J5b" > agent/state/J5_task2.id; echo "J5 task2: $J5b"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -14
