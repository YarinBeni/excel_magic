# TIMEOUT=300
# J15/J15b scored every rule with the DEFAULT backbone (TabPFN v2): the run config is nested under "config" in config.json
# and the script fell back to "tabpfn". Fixed (and the runner now seeds torch per evaluation). Cancel J15b, run the three
# groups in parallel as J15c.
id=$(cat agent/state/J15b.id 2>/dev/null || true); [ -n "$id" ] && scancel "$id" && echo "cancelled J15b $id"
J15c=$(sbatch --parsable sbatch/J15c_selection_rules.sbatch); echo "$J15c" > agent/state/J15c.id; echo "J15c: $J15c"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
