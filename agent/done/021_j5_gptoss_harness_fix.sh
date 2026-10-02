# TIMEOUT=300
# J5 task 3 (gpt-oss-20b, 50425): vLLM served fine this time, but the model read a .parquet with read_file and the
# UnicodeDecodeError aborted every search. Tool errors now return to the model; resubmit array task 3 only.
J5d=$(sbatch --parsable --array=3 sbatch/J5_llm_sweep.sbatch); echo "$J5d" > agent/state/J5_task3.id; echo "J5 task3: $J5d"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -14
