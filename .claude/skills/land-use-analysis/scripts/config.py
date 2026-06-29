"""Run configuration loader. Every deterministic script is driven by a single
per-run config.json (the locale-specific knowledge Claude discovers and authors:
subject point, analysis area, parcel/zoning endpoints + field maps, land-use code
scheme, and the zone->MF-threat crosswalk). See references/config-schema.md.

Resolution order for the config path:
  1. --config <path> on the command line
  2. $LUA_CONFIG
  3. ./config.json in the run directory ($LUA_ROOT or CWD)

The run directory (where in/, Tables/, and deliverables are written) is $LUA_ROOT
or the directory containing the config, falling back to CWD.
"""
import json
import os
import sys
from pathlib import Path


def _arg(flag):
    if flag in sys.argv:
        i = sys.argv.index(flag)
        if i + 1 < len(sys.argv):
            return sys.argv[i + 1]
    return None


def config_path():
    p = _arg("--config") or os.environ.get("LUA_CONFIG")
    if p:
        return Path(p).resolve()
    root = os.environ.get("LUA_ROOT")
    base = Path(root) if root else Path.cwd()
    return (base / "config.json").resolve()


def load():
    cp = config_path()
    if not cp.exists():
        sys.exit(f"[config] not found: {cp}\n"
                 "Author a config.json for this property first "
                 "(see references/config-schema.md + templates/config.example.json).")
    cfg = json.loads(cp.read_text(encoding="utf-8"))
    cfg["_path"] = str(cp)
    cfg["_root"] = str(_arg("--root")
                       and Path(_arg("--root")).resolve()
                       or os.environ.get("LUA_ROOT")
                       and Path(os.environ["LUA_ROOT"]).resolve()
                       or cp.parent)
    return cfg


def root(cfg):
    return Path(cfg["_root"])


def indir(cfg):
    d = root(cfg) / "in"
    d.mkdir(parents=True, exist_ok=True)
    return d


def tables(cfg):
    d = root(cfg) / "Tables"
    d.mkdir(parents=True, exist_ok=True)
    return d
