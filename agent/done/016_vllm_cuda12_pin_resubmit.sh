# TIMEOUT=300
# J0v 50342 installed vllm 0.30.0 -> torch 2.13+cu130, but the node driver is CUDA 12.8 ("driver too old", cuda unavailable),
# so every vLLM server in J3/J5 died at engine start. Installer now pins vllm 0.16.0 (torch 2.9.1 cu128), fallback 0.11.0,
# and asserts a CUDA 12 torch build. Wipe the env marker, rerun J0v, re-chain J3/J5/J6.
for k in J3 J5 J6; do id=$(cat agent/state/$k.id 2>/dev/null || true); [ -n "$id" ] && scancel "$id" 2>/dev/null && echo "cancelled $k $id"; done
rm -f "$HOME/miniconda3/envs/vllm/.ok"
sleep 3
J0v=$(sbatch --parsable sbatch/J0v_vllm.sbatch); echo "$J0v" > agent/state/J0v.id; echo "J0v: $J0v"
J3=$(sbatch --parsable --dependency=afterok:$J0v sbatch/J3_vllm_search.sbatch); echo "$J3" > agent/state/J3.id; echo "J3: $J3"
J5=$(sbatch --parsable --dependency=afterok:$J0v sbatch/J5_llm_sweep.sbatch); echo "$J5" > agent/state/J5.id; echo "J5: $J5"
J6=$(sbatch --parsable --dependency=afterok:$J0v sbatch/J6_cli_agents.sbatch); echo "$J6" > agent/state/J6.id; echo "J6: $J6"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -16
