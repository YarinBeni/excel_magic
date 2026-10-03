# TIMEOUT=300
# J20 ran the Wikidata triplet task (30/86 jobs); the entity-matching datasets were not built because the creation script
# asserts its raw-data folder exists before downloading. Folder now created; rerun row similarity search only.
J20b=$(sbatch --parsable --export=ALL,RUN_CONFIG=tabpfn_ctx_rowsim sbatch/J20_tembed_rows.sbatch); echo "$J20b" > agent/state/J20.id; echo "J20b: $J20b"
