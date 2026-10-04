# TIMEOUT=120
cd ~/sota
echo "######## RelAgent table"; sed -n 333,362p 2605.07840.txt | cut -c1-230
echo "######## ReViSQL"; grep -a -n -i -E "Arcwise-Plat-Full|Arcwise-Plat-SQL|human|Qwen3-Coder|self-consistency" 2603.20004.txt | head -40 | cut -c1-230
echo "######## InsightBench"; grep -a -n -i -E "G-Eval|LLaMA-3-Eval|ROUGE|AgentPoirot|Pandas Agent|0\.[0-9][0-9] *\(|insight-level|summary-level" 2407.06423.txt | head -45 | cut -c1-230
