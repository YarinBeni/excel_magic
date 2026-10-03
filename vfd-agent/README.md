# vfd-agent: Verified, Fast, Deep

A three-model harness for data-analysis agents: a reasoning LLM plans and writes SQL, GLiNER / GLiClass link the question
to the schema and triage each step and claim, and frozen tabular / relational foundation models (Kumo Tabular S/L,
Kumo Relational, TabPFN) answer what neither can: predictions, drivers, what-ifs, anomalies, drift, similar entities.
Plan and research: `../docs/research_3models/`.

Phase P0 (this code): how good are cheap signals at telling a correct SQL candidate from a wrong one, and at finding
the columns a question needs? `experiments/p0_verifier_study.py`, cluster job `sbatch/V1_verifier_study.sbatch`.
