# TIMEOUT=300
mkdir -p ~/sota/blog && cd ~/sota/blog
curl -sSL -m 60 -A "Mozilla/5.0" -o blog.html "https://www.sean-weldon.com/blog/2026-05-22-harnesses-in-ai-a-deep-dive-tejas-kumar-ibm"; ls -la blog.html
python3 - <<'P'
import re, html
s=open("blog.html",encoding="utf-8",errors="ignore").read()
print("LINKS:", sorted(set(re.findall(r'https?://(?:www\.)?(?:youtube\.com/watch\?v=[\w-]+|youtu\.be/[\w-]+|[^"\'<> ]*arxiv[^"\'<> ]*|github\.com/[^"\'<> ]+)', s)))[:30])
s=re.sub(r'(?is)<(script|style|nav|footer|head).*?</\1>','',s)
t=html.unescape(re.sub(r'<[^>]+>','\n',s)); t=re.sub(r'\n\s*\n+','\n',t)
open("blog.txt","w").write(t); print(len(t)); print(t[:12000])
P
