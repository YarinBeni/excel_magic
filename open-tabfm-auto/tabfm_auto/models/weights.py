"""Weight resolution for every backbone, with one cache directory for all of them.

Every Hugging Face based library (tabpfn, tabicl, sdm, exaonetabular, huggingface_hub) is pointed at
``weights/hf-cache`` so that:
  * once huggingface.co is reachable, ``tabfm-models download <name>`` fetches into the repo's weights dir;
  * on a machine without HF access the same dir can be rsynced in from a machine that has it (e.g. the cluster);
  * ``tabfm-models status`` tells you what is runnable right now.
"""
from __future__ import annotations

import argparse
import os
import socket
import sys
import urllib.request
from pathlib import Path

from .manifest import MODELS, ModelCard, find


def _default_weights_dir() -> Path:
    """``$TABFM_WEIGHTS_DIR`` if set; else ``./weights`` when it exists (repo checkout); else ``~/.cache/tabfm-auto/weights``."""
    if os.environ.get("TABFM_WEIGHTS_DIR"):
        return Path(os.environ["TABFM_WEIGHTS_DIR"])
    local = Path.cwd() / "weights"
    if local.is_dir():
        return local
    return Path.home() / ".cache" / "tabfm-auto" / "weights"


WEIGHTS_DIR = _default_weights_dir()
# child processes (tabfm-eval sandbox, agent shells) may run with another cwd: pin the resolved dir for them
os.environ.setdefault("TABFM_WEIGHTS_DIR", str(WEIGHTS_DIR))
HF_CACHE = WEIGHTS_DIR / "hf-cache"

# must run before any library imports huggingface_hub
os.environ.setdefault("HF_HOME", str(HF_CACHE))
os.environ.setdefault("HF_HUB_CACHE", str(HF_CACHE / "hub"))
os.environ.setdefault("TABPFN_MODEL_CACHE_DIR", str(WEIGHTS_DIR / "tabpfn"))


def hf_reachable(timeout: float = 5.0) -> bool:
    try:
        req = urllib.request.Request("https://huggingface.co/api/models/Prior-Labs/TabPFN-v2-clf", method="HEAD")
        urllib.request.urlopen(req, timeout=timeout)
        return True
    except Exception:
        return False


def package_installed(pkg: str) -> bool:
    import importlib.util

    for tp in (Path.cwd() / "third_party", WEIGHTS_DIR.parent / "third_party"):
        if tp.is_dir() and str(tp) not in sys.path:
            sys.path.insert(0, str(tp))  # vendored packages (e.g. OpenRFM in the research repo)
    return importlib.util.find_spec(pkg.split(" ")[0]) is not None


def local_paths(card: ModelCard) -> list[Path]:
    return [WEIGHTS_DIR / p for p in card.local_files]


def in_hf_cache(card: ModelCard) -> bool:
    if not card.hf_repo:
        return False
    try:
        from huggingface_hub import try_to_load_from_cache
    except ImportError:
        return False
    if not card.hf_files:  # whole-snapshot models: check the repo folder exists
        return any((HF_CACHE / "hub").glob(f"models--{card.hf_repo.replace('/', '--')}"))
    return all(isinstance(try_to_load_from_cache(card.hf_repo, f, cache_dir=HF_CACHE / "hub",
                                                 revision=card.hf_revision), str) for f in card.hf_files)


def status(name: str) -> dict:
    card = find(name)
    pkg_ok = package_installed(card.package)
    local = bool(card.local_files) and all(p.exists() and p.stat().st_size > 1000 for p in local_paths(card))
    cached = in_hf_cache(card)
    if local or cached:
        state = "ready"
    elif card.mirror_urls:
        state = "downloadable (mirror)"
    elif card.hf_repo:
        state = "needs huggingface.co"
    else:
        state = "manual"
    return {"name": card.name, "kind": card.kind, "state": state if pkg_ok else f"package {card.package} missing",
            "package_ok": pkg_ok, "local": local, "hf_cached": cached, "elo": card.elo_tabarena,
            "license": card.license, "cpu": card.cpu_ok, "params_m": card.params_m}


def available(name: str) -> bool:
    return status(name)["state"] == "ready"


def download(name: str, prefer_mirror: bool = True) -> list[Path]:
    """Fetch a model's weights into weights/. Mirror first (works without HF), then Hugging Face."""
    card = find(name)
    out: list[Path] = []
    if card.mirror_urls and prefer_mirror:
        for url, rel in zip(card.mirror_urls, card.local_files):
            dst = WEIGHTS_DIR / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            if not (dst.exists() and dst.stat().st_size > 1000):
                print(f"  {url} -> {dst}")
                urllib.request.urlretrieve(url, dst)
            out.append(dst)
        return out
    if not card.hf_repo:
        raise RuntimeError(f"{name}: no download source; see manifest notes")
    if not hf_reachable():
        raise RuntimeError(f"{name}: huggingface.co is not reachable from this machine (network policy). "
                           "Ask for the host to be allowed, or rsync weights/ from a machine that has access.")
    from huggingface_hub import hf_hub_download, snapshot_download

    if card.hf_files:
        for f in card.hf_files:
            p = hf_hub_download(card.hf_repo, f, revision=card.hf_revision, cache_dir=HF_CACHE / "hub")
            out.append(Path(p))
            # tabpfn / tabicl also accept a plain local path; mirror the file into weights/<pkg>/ for them
            if card.package in ("tabpfn", "tabicl"):
                dst = WEIGHTS_DIR / card.package / Path(f).name
                dst.parent.mkdir(parents=True, exist_ok=True)
                if not dst.exists():
                    os.symlink(p, dst)
    else:
        out.append(Path(snapshot_download(card.hf_repo, revision=card.hf_revision, cache_dir=HF_CACHE / "hub")))
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="tabfm-models", description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("status", help="what is runnable right now")
    d = sub.add_parser("download", help="fetch weights for one model, or 'all'")
    d.add_argument("names", nargs="+")
    a = ap.parse_args(argv)
    if a.cmd == "status":
        hf = hf_reachable()
        print(f"weights dir: {WEIGHTS_DIR}   huggingface.co reachable: {hf}   host: {socket.gethostname()}")
        print(f"| {'model':16s} | {'kind':10s} | {'state':24s} | {'TabArena Elo':34s} | {'licence':30s} | cpu |")
        print("|---|---|---|---|---|---|")
        for n in MODELS:
            s = status(n)
            print(f"| {s['name']:16s} | {s['kind']:10s} | {s['state']:24s} | {s['elo']:34s} | {s['license'][:30]:30s} | {s['cpu']} |")
        return 0
    names = list(MODELS) if a.names == ["all"] else a.names
    rc = 0
    for n in names:
        try:
            print(f"{n}:")
            for p in download(n):
                print(f"  ok {p}")
        except Exception as e:
            print(f"  FAILED: {e}")
            rc = 1
    return rc


if __name__ == "__main__":
    sys.exit(main())
