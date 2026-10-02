weights dir: /home/yarin.b/projects/excel_magic/artifacts/weights   huggingface.co reachable: True   host: g0381
| model            | kind       | state                    | TabArena Elo                       | licence                        | cpu |
|---|---|---|---|---|---|
| tabpfn           | tabular    | ready                    | not in pool; below RealTabPFN-2.5 (1522) | Prior Labs License (Apache-2.0 | yes (<=10k rows) |
| tabpfn-2.5       | tabular    | ready                    | 1522 (RealTabPFN-2.5)              | TabPFN-2.5 license (non-commer | yes |
| tabpfn-2.6       | tabular    | needs huggingface.co     | 1605                               | TabPFN-2.6 license (non-commer | yes |
| tabpfn-3         | tabular    | needs huggingface.co     | 1660                               | non-commercial                 | slow (default limit 5000 rows) |
| tabpfn-3.5       | tabular    | needs huggingface.co     | ~1855                              | academic + evaluation          | slow |
| tabpfn-3.5-fast  | tabular    | needs huggingface.co     | ~1770                              | academic + evaluation          | probably |
| tabicl           | tabular    | ready                    | 1590                               | BSD-3                          | yes (GPU recommended) |
| tabiclv2-sdm     | tabular    | needs huggingface.co     | 1590                               | BSD-3                          | yes |
| kumo-tabular-s   | tabular    | ready                    | ~1790                              | OpenMDW-1.1 (commercial OK)    | expected (TabICLv2-size) |
| kumo-tabular-m   | tabular    | ready                    | ~1910                              | OpenMDW-1.1                    | slow |
| kumo-tabular-l   | tabular    | ready                    | ~1960 (#1)                         | OpenMDW-1.1                    | impractical |
| exaone           | tabular    | ready                    | 1765                               | code BSD-3; weights non-commer | slow |
| openrfm          | relational | package kumorfm_repro (vendored) missing | RelBench 12-task AUROC ~69 (KumoRFM-2: 79.6) | MIT                            | yes |
| kumo-relational  | relational | ready                    | KumoRFM-2 adapted                  | OpenMDW-1.1                    | expected |
| rt-j             | relational | package rt missing       | RelBench zero-shot 93% of supervised | see repo                       | eager flex_attention |
