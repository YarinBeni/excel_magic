# TIMEOUT=600
source "$HOME/venvs/vfd/bin/activate"
R=reports/runs/20261003T202238Z_V5_sql_harness
python vfd-agent/experiments/analyze_failures.py --rows "$R/harness_rows.jsonl" --config A --db-root artifacts/bird --cache artifacts/vfd_data 2>&1 | cut -c1-400
python vfd-agent/experiments/analyze_failures.py --rows "$R/harness_rows.jsonl" --config SC --db-root artifacts/bird --cache artifacts/vfd_data 2>&1 | head -12
