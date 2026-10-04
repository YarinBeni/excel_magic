# TIMEOUT=120
cd ~/sota
sed -n 700,770p 2603.20004.txt | cut -c1-210
echo "#### InsightBench numbers"; grep -a -n -E "0\.[0-9]{2} ?± ?0\.[0-9]{2}|0\.[0-9]{2} \(0\.[0-9]{2}\)" 2407.06423.txt | head -30 | cut -c1-200
