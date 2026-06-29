"""Orchestrate the deterministic pipeline. Two intentional stops:
  1. ALLOWLIST gate — after authoring config.json, `endpoints` lists every domain the run
     will contact. Pulls refuse to run until the user has allowlisted them and the flag file
     in/allowlist_ok exists (sandboxed environments block un-allowlisted egress — the usual
     "it doesn't work" cause).
  2. REASON gate — after `prep`, you must reason over the candidates and write
     in/reasoned_ranking.json (see references/threat-reasoning.md) before `finalize`/`deliver`.

NOTE: allowlisting a domain only takes effect in a NEW session (the sandbox reads the network
policy once, at session start). And allowlisted != reachable — `pull` runs a reachability
pre-flight first and halts if a host is blocked (allowlist) or unreachable (503/timeout origin).

Usage (from the run directory, with config.json present):
  python <skill>/scripts/run_all.py endpoints   # list domains to allowlist + next-session prompt
  python <skill>/scripts/run_all.py preflight    # probe each endpoint (allowlisted != reachable)
  python <skill>/scripts/run_all.py pull         # preflight, then parcels + zoning (needs in/allowlist_ok)
  python <skill>/scripts/run_all.py classify     # land-use + zoning + threat join
  python <skill>/scripts/run_all.py prep         # vacant candidates (then REASON)
  python <skill>/scripts/run_all.py finalize     # needs in/reasoned_ranking.json
  python <skill>/scripts/run_all.py deliver      # viewer + workbook
  python <skill>/scripts/run_all.py all          # endpoints->prep->(stops at each gate)
Pass --config / --root through to all stages (forwarded automatically).
"""
import runpy
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import config as C  # noqa: E402


def run(mod):
    print(f"\n=== {mod} ===")
    runpy.run_path(str(HERE / f"{mod}.py"), run_name="__main__")


def main():
    stage = next((a for a in sys.argv[1:] if not a.startswith("-")), "all")
    cfg = C.load()
    IN = C.indir(cfg)
    allowlist_ok = IN / "allowlist_ok"
    ranking = IN / "reasoned_ranking.json"

    if stage in ("endpoints", "all"):
        run("list_endpoints")
    if stage == "all" and not allowlist_ok.exists():
        print("\n*** STOP (allowlist gate): have the user allowlist the domains above, then "
              "create the flag file in/allowlist_ok and run: run_all.py pull ***")
        return

    if stage in ("preflight", "pull", "all"):
        if not allowlist_ok.exists():
            sys.exit("[allowlist] in/allowlist_ok not found. Run `endpoints`, get the domains "
                     "allowlisted, then `touch in/allowlist_ok` (or write any content) and retry.")
        run("preflight")  # exits nonzero (halting the run) if any source is blocked/unreachable
    if stage in ("pull", "all"):
        run("pull_parcels"); run("pull_zoning")
    if stage in ("classify", "all"):
        run("classify")
    if stage in ("prep", "all"):
        run("prepare_candidates")
    if stage == "all" and not ranking.exists():
        print("\n*** STOP (reason gate): reason over in/vacant_candidates.json and write "
              "in/reasoned_ranking.json, then run: run_all.py deliver ***")
        return
    if stage in ("finalize", "deliver", "all"):
        run("finalize_topN")
    if stage in ("deliver", "all"):
        run("build_viewer"); run("build_workbook")
        print("\nDone. Deliverables written to the run directory.")


if __name__ == "__main__":
    main()
