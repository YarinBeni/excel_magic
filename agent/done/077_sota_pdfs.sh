# TIMEOUT=300
mkdir -p ~/sota && cd ~/sota
for id in 2407.20060 2604.12596 2605.07840 2407.06423 2603.20004 2111.02705; do
  [ -s $id.pdf ] || curl -sSL -m 60 -o $id.pdf "https://arxiv.org/pdf/$id" ; done
ls -la
PY=$(command -v python3); for v in ~/venvs/vfd/bin/python ~/miniconda3/envs/vllm/bin/python; do [ -x $v ] && PY=$v; done
$PY - <<'P'
import re, glob
try:
    from pypdf import PdfReader
except ImportError:
    import subprocess; subprocess.run(["pip","install","-q","pypdf"]); from pypdf import PdfReader
for f in sorted(glob.glob("*.pdf")):
    try: t="\n".join((p.extract_text() or "") for p in PdfReader(f).pages)
    except Exception as e: print("==",f,"ERR",e); continue
    open(f[:-4]+".txt","w").write(t); print("==",f,len(t))
P
