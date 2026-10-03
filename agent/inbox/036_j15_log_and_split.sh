# TIMEOUT=300
# J15 takes ~3 h per group (5 rules x 30 splits x Kumo-S fits; the ensemble = 3 fits per split) and will hit its 8 h limit in
# the middle of the GLM group. Show its log, cancel it, and run the GLM group alone (J15b); the budget-64 groups are skipped.
f=$(ls -t sbatch/logs/J15_*.out | head -1); echo "== $f"; grep -a "\[rescore\]" "$f" | sed 's/picks=.*} //' | tail -40 | cut -c1-150
echo "== slowest"; grep -a "\[rescore\]" "$f" | sed 's/.*Z_//; s/ picks=.*} / /' | sort -k3 -n -r | head -8
id=$(cat agent/state/J15.id 2>/dev/null || true); [ -n "$id" ] && scancel "$id" && echo "cancelled J15 $id"
J15b=$(sbatch --parsable sbatch/J15b_selection_glm.sbatch); echo "$J15b" > agent/state/J15b.id; echo "J15b: $J15b"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
