# TIMEOUT=120
cd ~/sota
echo "######## KumoRFM-2 Table 3"; sed -n 520,560p 2604.12596.txt | cut -c1-260
echo "######## RelAgent"; grep -a -n -i -E "driver-dnf|user-churn|dnf|top3|Avg|backbone|GPT-|Claude|Qwen|gemini" 2605.07840.txt | head -40 | cut -c1-260
