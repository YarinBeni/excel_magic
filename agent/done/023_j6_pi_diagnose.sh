# TIMEOUT=300
# J6 task 0 (pi) left the queue with no results commit and no log push. Show what happened.
sacct -j 50377 -o JobID,JobName%16,State,ExitCode,Elapsed,MaxRSS,NodeList --parsable2 2>/dev/null | head -12
echo "== artifacts"; ls -d artifacts/runs/*J6_pi* 2>/dev/null | head
echo "== sbatch/logs/J6_cli_agents_50377_0.out (filtered tail)"
grep -avE "^\s*$|it/s\]|%\|" sbatch/logs/J6_cli_agents_50377_0.out 2>/dev/null | grep -av "APIServer pid" | tail -60
echo "== pi binary"; export PATH="$HOME/miniconda3/envs/node22/bin:$PATH"; command -v pi && pi --version 2>&1 | head -2; ls ~/.pi/agent 2>/dev/null
