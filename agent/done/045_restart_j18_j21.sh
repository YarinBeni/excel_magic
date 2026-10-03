# TIMEOUT=300
# J18 hooked whichever sdm network came first in a set-ordered dict (regression in some processes) -> no layers recorded;
# fixed to the classification network, smoke now asserts the layers exist. J21's smoke flagged only the unscaled raw
# features (not a model layer); sdm models are now built directly. Restart J18, resubmit J21.
id=$(cat agent/state/J18.id 2>/dev/null || true); [ -n "$id" ] && scancel "$id" && echo "cancelled J18 $id"
sleep 3
J18=$(sbatch --parsable sbatch/J18_relbench_layers.sbatch); echo "$J18" > agent/state/J18.id; echo "J18: $J18"
J21=$(sbatch --parsable sbatch/J21_model_layers.sbatch); echo "$J21" > agent/state/J21.id; echo "J21: $J21"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
