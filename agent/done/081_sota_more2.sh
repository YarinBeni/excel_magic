# TIMEOUT=120
cd ~/sota
echo "######## InsightBench results table"; sed -n 405,470p 2407.06423.txt | cut -c1-200
echo "######## ReViSQL results table"; grep -a -n -E "^ *Table [0-9]" 2603.20004.txt | head -20 | cut -c1-200
n=$(grep -a -n -i "Table 4\|Table 3:" 2603.20004.txt | head -1 | cut -d: -f1); sed -n "$((n)),$((n+45))p" 2603.20004.txt | cut -c1-200
