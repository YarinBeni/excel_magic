# TIMEOUT=300
# J0 installed everything except NVIDIA sdm (wrong pip name) -> Kumo backbones missing. Cancel the pending chain,
# re-run J0 (idempotent: only adds sdm + symlinks), resubmit J1..J6 after it.
for id in $(cat agent/state/J1.id agent/state/J4.id agent/state/J2.id agent/state/J3.id agent/state/J5.id agent/state/J6.id 2>/dev/null); do scancel "$id" 2>/dev/null && echo "cancelled $id"; done
sleep 3
J0=$(sbatch --parsable sbatch/J0_setup.sbatch); echo "$J0" > agent/state/J0.id; echo "J0b: $J0"
J1=$(sbatch --parsable --dependency=afterok:$J0 sbatch/J1_smoke.sbatch); echo "$J1" > agent/state/J1.id; echo "J1: $J1"
J4=$(sbatch --parsable --dependency=afterok:$J0 sbatch/J4_relbench_probe.sbatch); echo "$J4" > agent/state/J4.id; echo "J4: $J4"
J2=$(sbatch --parsable --dependency=afterok:$J0 sbatch/J2_tabarena.sbatch); echo "$J2" > agent/state/J2.id; echo "J2: $J2"
J3=$(sbatch --parsable --dependency=afterok:$J0 sbatch/J3_vllm_search.sbatch); echo "$J3" > agent/state/J3.id; echo "J3: $J3"
J5=$(sbatch --parsable --dependency=afterok:$J0 sbatch/J5_llm_sweep.sbatch); echo "$J5" > agent/state/J5.id; echo "J5: $J5"
J6=$(sbatch --parsable --dependency=afterok:$J0 sbatch/J6_cli_agents.sbatch); echo "$J6" > agent/state/J6.id; echo "J6: $J6"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -12
