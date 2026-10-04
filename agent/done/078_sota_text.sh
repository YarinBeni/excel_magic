# TIMEOUT=300
cd ~/sota
if command -v pdftotext >/dev/null; then for f in *.pdf; do pdftotext -layout $f ${f%.pdf}.txt; done; echo pdftotext
else PY=$HOME/venvs/vfd/bin/python; $PY -m pip install -q pypdf 2>&1 | tail -1
  $PY -c "
import glob
from pypdf import PdfReader
for f in glob.glob('*.pdf'):
    open(f[:-4]+'.txt','w').write('\n'.join((p.extract_text(extraction_mode='layout') or '') for p in PdfReader(f).pages))
"; fi
wc -c *.txt
echo "######## RelBench (2407.20060) entity classification"
grep -n -i -E "user-churn|driver-dnf|driver-top3|study-outcome|user-repeat|user-ignore|user-visits|user-clicks" 2407.20060.txt | head -40
echo "######## KumoRFM-2 (2604.12596)"
grep -n -i -E "user-churn|driver-dnf|driver-top3|study-outcome|user-repeat|user-ignore|user-visits|user-clicks|Avg|Average" 2604.12596.txt | head -60
echo "######## RelAgent (2605.07840)"
grep -n -i -E "user-churn|driver-dnf|driver-top3|study-outcome|user-repeat|user-ignore|user-visits|user-clicks|Avg|Average|GPT|Claude|Qwen" 2605.07840.txt | head -60
