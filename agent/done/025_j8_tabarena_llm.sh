# TIMEOUT=300
# J8: the paper's TabArena protocol (17 datasets) with the two best free setups: pi + Qwen3-Coder (task 0) and our tool
# loop + GLM-4.5-Air-FP8 (task 1). One GPU each, results pushed per dataset.
if squeue -u "$USER" -h -o "%j" | grep -q '^J8_tabarena_llm$'; then echo "J8 already queued"; exit 0; fi
J8=$(sbatch --parsable sbatch/J8_tabarena_llm.sbatch); echo "$J8" > agent/state/J8.id; echo "J8: $J8"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -8
