#!/bin/bash
#SBATCH --job-name=J1_smoke
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=64G
#SBATCH --time=04:00:00
#SBATCH --output=sbatch/logs/J1_smoke_%j.out
# J1 — first real results on the cluster (GPU):
#   a) backbone comparison with the identity pipeline on the bundled datasets (TabPFN v2 vs TabICLv2 vs Kumo Tabular-S vs HistGB)
#   b) no-LLM heuristic search around Kumo Tabular-S on the synthetic datasets
#   c) backbone transfer of the pipelines found so far
#   d) retrieval benchmark, 3 seeds, incl. KumoRelational hidden states (first time)
set -uo pipefail
echo "=== J1 smoke === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd open-tabfm-auto
tabfm-models download kumo-tabular-m kumo-tabular-l 2>&1 | tail -3 || true
tabfm-models status > ../reports/J1_models_status.md
python examples/01_tfm_vs_baselines.py --name J1_backbones \
    --models "tabpfn:n_estimators=4,device=cuda,tabpfn-2.5:n_estimators=4,device=cuda,tabicl:device=cuda,kumo-tabular-s:n_estimators=4,device=cuda,kumo-tabular-m:n_estimators=4,device=cuda,kumo-tabular-l:n_estimators=4,device=cuda,exaone:device=cuda,hgb" \
    --datasets wine,breast_cancer,diabetes,synth_physics,synth_entities || echo "[J1] a) FAILED"
for ds in synth_physics synth_entities; do
    tabfm-auto search --dataset $ds --harness heuristic --model "kumo-tabular-s:n_estimators=4,device=cuda" \
        --budget-evals 20 --budget-minutes 30 --baselines hgb --run-name "J1_heur_kumo_$ds" || echo "[J1] b) $ds FAILED"
done
python examples/04_backbone_transfer.py --models "tabpfn,tabpfn-2.5,tabicl,kumo-tabular-s,kumo-tabular-m,kumo-tabular-l,exaone,hgb" --name J1_transfer || echo "[J1] c) FAILED"
cd ../frozen-embeddings-retrieval
for s in 0 1 2; do
    python experiments/exp01_retrieval_benchmark.py --seed $s --name "J1_retrieval_seed$s" \
        --embedders row,agg,tabpfn_agg_kmeans,tabpfn_agg,gnn,openrfm,kumo_relational || echo "[J1] d) seed $s FAILED"
done
python scripts/rescore.py >/dev/null 2>&1 || true
python scripts/aggregate_runs.py > ../reports/J1_retrieval_aggregate.md 2>&1 || true
cd ..
collect_and_push "J1: smoke results (job ${SLURM_JOB_ID})"
echo "=== J1 done $(date) ==="
