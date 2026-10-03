# TIMEOUT=300
# J17: the recommended free configuration end to end (Kumo-L, cv-repeats auto, gated1 pick, 20% acceptance slice),
# heuristic (task 0) and pi + Qwen3-Coder (task 1), TabArena wave 1.
if squeue -u "$USER" -h -o "%j" | grep -q '^J17_recommended$'; then echo "J17 already queued"; exit 0; fi
J17=$(sbatch --parsable sbatch/J17_recommended.sbatch); echo "$J17" > agent/state/J17.id; echo "J17: $J17"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
